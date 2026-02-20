#!/usr/bin/env python3
"""
Cap Quest Master Report
Generates tabs: Master Dashboard, Campaigns 2026, Agents

Features:
- Tabs: Master Dashboard (1st), Campaigns 2026, Agents
- Master Dashboard: All-Time Data (2024+), Merged Headers (Left Aligned)
- Campaigns 2026: Weekly Total in Col A (Left Aligned), Wide Col B, Frozen Header, Grand Total YTD
- Email: Luxvance Branding (No Logo), Simplified Footer
"""

import os
import sys
import datetime
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from collections import defaultdict
import calendar

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from app.instantly_client import InstantlyClient
from app.email_reporter import send_email_report

load_dotenv()

# Configuration
SHEET_ID = os.getenv("CONNECT_RESOURCES_SHEET_ID")
CREDS_FILE = 'credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Luxvance Color Palette (Black & Gold)
COLORS = {
    'header_dark': {'red': 0, 'green': 0, 'blue': 0},               # #000000 (Black)
    'header_text_gold': {'red': 0.83, 'green': 0.68, 'blue': 0.21}, # #D4AF37 (Metallic Gold)
    'header_text_white': {'red': 1, 'green': 1, 'blue': 1},         # #FFFFFF
    'subheader_gray': {'red': 0.2, 'green': 0.2, 'blue': 0.2},      # #333333 (Dark Gray)
    'week_header_gold': {'red': 1.0, 'green': 0.95, 'blue': 0.8},   # #FFF2CC (Light Gold)
    'row_alt_gray': {'red': 0.98, 'green': 0.98, 'blue': 0.98},     # #FAFAFA
    'total_gold': {'red': 0.96, 'green': 0.87, 'blue': 0.70},       # #F4DEB3 (Wheat/Goldish)
    'grand_total_black': {'red': 0.1, 'green': 0.1, 'blue': 0.1},   # #1A1A1A
    'text_black': {'red': 0, 'green': 0, 'blue': 0},
    'bg_white': {'red': 1, 'green': 1, 'blue': 1}
}

def get_week_info(date_obj):
    """Get ISO week number and date range for a given date"""
    iso_year, iso_week, _ = date_obj.isocalendar()
    monday = date_obj - datetime.timedelta(days=date_obj.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return {
        'year': iso_year,
        'week_num': iso_week,
        'start': monday,
        'end': sunday,
        'label': f"Week {iso_week} (Monday, {monday.strftime('%b %d')} - Sunday, {sunday.strftime('%b %d, %Y')})"
    }

def fetch_all_data(client):
    """Fetch all data starting from 2024 to capture 'All Time'"""
    print("\n📊 Fetching all data (from 2024-01-01)...")
    
    start_date = datetime.date(2024, 1, 1)
    end_date = datetime.date.today()
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')
    
    # Fetch campaigns
    print("  → Fetching campaigns...")
    campaigns = []
    skip = 0
    while True:
        resp = client.request("GET", "/campaigns", params={"limit": 100, "skip": skip})
        items = resp.get('items', []) if isinstance(resp, dict) else resp if isinstance(resp, list) else []
        if not items:
            break
        campaigns.extend(items)
        if len(items) < 100:
            break
        skip += 100
    print(f"  ✓ Found {len(campaigns)} campaigns")
    
    # Fetch campaign daily analytics
    print("  → Fetching campaign analytics...")
    campaign_analytics = {}
    for camp in campaigns:
        c_id = camp.get('id')
        daily_stats = client.request("GET", "/campaigns/analytics/daily", params={
            "campaign_id": c_id,
            "start_date": start_str,
            "end_date": end_str
        })
        if daily_stats and isinstance(daily_stats, list):
            campaign_analytics[c_id] = daily_stats
    print(f"  ✓ Fetched analytics for {len(campaign_analytics)} campaigns")
    
    # Fetch accounts
    print("  → Fetching accounts...")
    accounts = []
    starting_after = None
    while True:
        params = {"limit": 100}
        if starting_after:
            params["starting_after"] = starting_after
        resp = client.request("GET", "/accounts", params=params)
        items = resp.get('items', []) if isinstance(resp, dict) else resp if isinstance(resp, list) else []
        if not items:
            break
        accounts.extend(items)
        if isinstance(resp, dict) and resp.get('next_starting_after'):
            starting_after = resp['next_starting_after']
        else:
            break
    print(f"  ✓ Found {len(accounts)} accounts")
    
    # Fetch account daily analytics
    print("  → Fetching account analytics...")
    account_analytics = client.request("GET", "/accounts/analytics/daily", params={
        "start_date": start_str,
        "end_date": end_str
    })
    print(f"  ✓ Fetched {len(account_analytics) if isinstance(account_analytics, list) else 0} daily records")
    
    return {
        'campaigns': campaigns,
        'campaign_analytics': campaign_analytics,
        'accounts': accounts,
        'account_analytics': account_analytics,
        'start_date': start_date,
        'end_date': end_date
    }

def create_master_dashboard(sh, data):
    """Create Master Dashboard (Tab 1) - All Time & Monthly"""
    tab_name = "Master Dashboard"
    print(f"\n📋 Creating {tab_name}...")
    
    # Aggregate by Year-Month
    year_data = defaultdict(lambda: defaultdict(lambda: {
        'sent': 0, 'new_leads': 0, 'replies': 0, 'opportunities': 0
    }))
    
    # All Time Totals
    all_time_totals = {'sent': 0, 'new_leads': 0, 'replies': 0, 'opportunities': 0}

    for camp in data['campaigns']:
        c_id = camp.get('id')
        daily_stats = data['campaign_analytics'].get(c_id, [])
        for day in daily_stats:
            date_str = day.get('date')
            if not date_str: continue
            
            date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
            year = date_obj.year
            month_key = date_obj.strftime('%Y-%m') # YYYY-MM
            
            # Monthly Stats
            year_data[year][month_key]['sent'] += day.get('sent', 0)
            year_data[year][month_key]['new_leads'] += day.get('new_leads_contacted', 0)
            year_data[year][month_key]['replies'] += day.get('unique_replies', 0)
            year_data[year][month_key]['opportunities'] += day.get('opportunities', 0)

            # All Time Stats
            all_time_totals['sent'] += day.get('sent', 0)
            all_time_totals['new_leads'] += day.get('new_leads_contacted', 0)
            all_time_totals['replies'] += day.get('unique_replies', 0)
            all_time_totals['opportunities'] += day.get('opportunities', 0)

    
    print(f"  ✓ Processed data for years: {list(year_data.keys())}")
    
    # Build rows
    rows = []
    
    # Header Section
    rows.append(['MASTER DASHBOARD', '', '', '', '', '']) # Row 1
    rows.append(['ALL TIME PERFORMANCE', '', '', '', '', '']) # Row 2
    rows.append(['Emails Sent', 'New Leads', 'Replies', 'Opportunities', '', '']) # Row 3
    rows.append([f"{all_time_totals['sent']:,}", f"{all_time_totals['new_leads']:,}", f"{all_time_totals['replies']:,}", f"{all_time_totals['opportunities']:,}", '', '']) # Row 4
    rows.append(['', '', '', '', '', '']) # Spacer
    
    # --- Performance by Month 2026 ---
    stats_2026 = year_data.get(2026, {})
    if stats_2026:
        rows.append(['PERFORMANCE BY MONTH 2026', '', '', '', '', ''])
        rows.append(['Month', 'Emails Sent', 'New Leads', 'Replies', 'Opportunities', 'Emails/Opp'])
        
        sorted_months_2026 = sorted(stats_2026.keys())
        for m in sorted_months_2026:
            stats = stats_2026[m]
            month_name = datetime.datetime.strptime(m, '%Y-%m').strftime('%B')
            emails_per_opp = stats['sent'] / stats['opportunities'] if stats['opportunities'] > 0 else 0
            rows.append([
                month_name,
                f"{stats['sent']:,}",
                f"{stats['new_leads']:,}",
                f"{stats['replies']:,}",
                f"{stats['opportunities']:,}",
                f"{emails_per_opp:.1f}"
            ])
        rows.append(['', '', '', '', '', '']) # Spacer

    # --- Performance by Month 2025 ---
    stats_2025 = year_data.get(2025, {})
    if stats_2025:
        rows.append(['PERFORMANCE BY MONTH 2025', '', '', '', '', ''])
        rows.append(['Month', 'Emails Sent', 'New Leads', 'Replies', 'Opportunities', 'Emails/Opp'])
        
        sorted_months_2025 = sorted(stats_2025.keys())
        for m in sorted_months_2025:
            stats = stats_2025[m]
            month_name = datetime.datetime.strptime(m, '%Y-%m').strftime('%B')
            emails_per_opp = stats['sent'] / stats['opportunities'] if stats['opportunities'] > 0 else 0
            rows.append([
                month_name,
                f"{stats['sent']:,}",
                f"{stats['new_leads']:,}",
                f"{stats['replies']:,}",
                f"{stats['opportunities']:,}",
                f"{emails_per_opp:.1f}"
            ])
            
    # Delete and recreate
    try:
        ws = sh.worksheet(tab_name)
        sh.del_worksheet(ws)
    except:
        pass
    
    ws = sh.add_worksheet(title=tab_name, rows=100, cols=10)
    
    # Write data
    ws.update(values=rows, range_name='A1')
    
    # Apply formatting
    print("  → Applying formatting...")
    requests = []

    # Merge Calls
    requests.append({
        'mergeCells': {
            'range': {'sheetId': ws.id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'mergeType': 'MERGE_ALL'
        }
    })
    requests.append({
        'mergeCells': {
            'range': {'sheetId': ws.id, 'startRowIndex': 1, 'endRowIndex': 2, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'mergeType': 'MERGE_ALL'
        }
    })
    
    sh.batch_update({'requests': requests})
    requests = [] # Reset for formatting
    
    # Helper to format a block
    def format_headers_req(row_idx):
        # Header (Merged or single)
        requests.append({
            'repeatCell': {
                'range': {'sheetId': ws.id, 'startRowIndex': row_idx, 'endRowIndex': row_idx+1, 'startColumnIndex': 0, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': COLORS['header_dark'],
                        'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_gold'], 'fontSize': 12},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
            }
        })
        # Column Header (Next Row)
        requests.append({
            'repeatCell': {
                'range': {'sheetId': ws.id, 'startRowIndex': row_idx+1, 'endRowIndex': row_idx+2, 'startColumnIndex': 0, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': COLORS['subheader_gray'],
                        'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_white']},
                        'horizontalAlignment': 'CENTER'
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
            }
        })

    # Row 1: Dashboard Title (LEFT Aligned)
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': COLORS['header_dark'],
                    'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_gold'], 'fontSize': 14},
                    'horizontalAlignment': 'LEFT',
                    'verticalAlignment': 'MIDDLE'
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
        }
    })
    
    # Row 2: All Time Header (LEFT Aligned)
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 1, 'endRowIndex': 2, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': COLORS['header_dark'],
                    'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_white'], 'fontSize': 12},
                    'horizontalAlignment': 'LEFT',
                    'verticalAlignment': 'MIDDLE'
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
        }
    })
    
    # Row 3: Column Headers
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 2, 'endRowIndex': 3, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': COLORS['subheader_gray'],
                    'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_white']},
                    'horizontalAlignment': 'CENTER'
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
        }
    })
    
    # Row 4: Values (Total)
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 3, 'endRowIndex': 4, 'startColumnIndex': 0, 'endColumnIndex': 6},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': COLORS['total_gold'], # Highlight total
                    'textFormat': {'bold': True, 'foregroundColor': COLORS['text_black'], 'fontSize': 11},
                    'horizontalAlignment': 'CENTER'
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
        }
    })
    
    # Find other tables
    for i, row in enumerate(rows):
        if row[0] == 'PERFORMANCE BY MONTH 2026':
            format_headers_req(i)
        elif row[0] == 'PERFORMANCE BY MONTH 2025':
            format_headers_req(i)
            
    sh.batch_update({'requests': requests})
    print(f"  ✓ {tab_name} styled")

def create_campaigns_tab(sh, data):
    """Create Campaigns 2026 (Tab 2)"""
    tab_name = "Campaigns 2026"
    print(f"\n📅 Creating {tab_name}...")
    
    target_year = 2026
    
    # Aggregate by (week, campaign_name)
    week_camp_data = defaultdict(lambda: {
        'sent': 0, 'new_leads': 0, 'replies': 0, 'opportunities': 0
    })
    
    for camp in data['campaigns']:
        c_id = camp.get('id')
        c_name = camp.get('name', 'Unnamed')
        
        daily_stats = data['campaign_analytics'].get(c_id, [])
        for day in daily_stats:
            date_str = day.get('date')
            if not date_str: continue
            
            date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
            if date_obj.year != target_year: continue
                
            week_info = get_week_info(date_obj)
            key = (week_info['week_num'], week_info['label'], c_name)
            week_camp_data[key]['sent'] += day.get('sent', 0)
            week_camp_data[key]['new_leads'] += day.get('new_leads_contacted', 0)
            week_camp_data[key]['replies'] += day.get('unique_replies', 0)
            week_camp_data[key]['opportunities'] += day.get('opportunities', 0)
    
    # Build rows
    rows = []
    rows.append(['Week', 'Campaign Name', 'Emails Sent', 'New Leads', 'Replies', 'Opportunities', 'Reply %', 'Opp %'])
    
    # Group by week
    weeks = {}
    for (week_num, week_label, c_name), stats in week_camp_data.items():
        if week_num not in weeks:
            weeks[week_num] = {'label': week_label, 'campaigns': {}}
        weeks[week_num]['campaigns'][c_name] = stats
    
    grand_totals = {'sent': 0, 'new_leads': 0, 'replies': 0, 'opportunities': 0}
    
    for week_num in sorted(weeks.keys()):
        week = weeks[week_num]
        
        # Week header
        rows.append([week['label'], '', '', '', '', '', '', ''])
        
        # Campaign rows
        week_totals = {'sent': 0, 'new_leads': 0, 'replies': 0, 'opportunities': 0}
        
        for c_name in sorted(week['campaigns'].keys()):
            stats = week['campaigns'][c_name]
            if stats['sent'] == 0 and stats['replies'] == 0 and stats['opportunities'] == 0: continue

            reply_pct = (stats['replies'] / stats['sent'] * 100) if stats['sent'] > 0 else 0
            opp_pct = (stats['opportunities'] / stats['sent'] * 100) if stats['sent'] > 0 else 0
            
            rows.append([
                '', c_name,
                f"{stats['sent']:,}", f"{stats['new_leads']:,}", f"{stats['replies']:,}", f"{stats['opportunities']:,}",
                f"{reply_pct:.1f}%", f"{opp_pct:.1f}%"
            ])
            
            week_totals['sent'] += stats['sent']
            week_totals['new_leads'] += stats['new_leads']
            week_totals['replies'] += stats['replies']
            week_totals['opportunities'] += stats['opportunities']
        
        # Weekly total
        week_reply_pct = (week_totals['replies'] / week_totals['sent'] * 100) if week_totals['sent'] > 0 else 0
        week_opp_pct = (week_totals['opportunities'] / week_totals['sent'] * 100) if week_totals['sent'] > 0 else 0
        
        rows.append([
            'WEEKLY TOTAL', '',
            f"{week_totals['sent']:,}", f"{week_totals['new_leads']:,}", f"{week_totals['replies']:,}", f"{week_totals['opportunities']:,}",
            f"{week_reply_pct:.1f}%", f"{week_opp_pct:.1f}%"
        ])
        rows.append(['', '', '', '', '', '', '', '']) # Spacer
        
        grand_totals['sent'] += week_totals['sent']
        grand_totals['new_leads'] += week_totals['new_leads']
        grand_totals['replies'] += week_totals['replies']
        grand_totals['opportunities'] += week_totals['opportunities']
    
    # Grand total
    grand_reply_pct = (grand_totals['replies'] / grand_totals['sent'] * 100) if grand_totals['sent'] > 0 else 0
    grand_opp_pct = (grand_totals['opportunities'] / grand_totals['sent'] * 100) if grand_totals['sent'] > 0 else 0
    
    rows.append([
        'GRAND TOTAL (YTD)', '',
        f"{grand_totals['sent']:,}", f"{grand_totals['new_leads']:,}", f"{grand_totals['replies']:,}", f"{grand_totals['opportunities']:,}",
        f"{grand_reply_pct:.1f}%", f"{grand_opp_pct:.1f}%"
    ])
    
    # Delete and recreate
    try:
        ws = sh.worksheet(tab_name)
        sh.del_worksheet(ws)
    except:
        pass
    
    ws = sh.add_worksheet(title=tab_name, rows=500, cols=10)
    ws.update(values=rows, range_name='A1')
    
    # Formatting
    print("  → Applying formatting...")
    requests = []
    
    # Freeze Header Row
    requests.append({
        'updateSheetProperties': {
            'properties': {'sheetId': ws.id, 'gridProperties': {'frozenRowCount': 1}},
            'fields': 'gridProperties.frozenRowCount'
        }
    })
    
    # Set Column Widths (Col B Wider)
    requests.append({
        'updateDimensionProperties': {
            'range': {'sheetId': ws.id, 'dimension': 'COLUMNS', 'startIndex': 1, 'endIndex': 2}, # Col B
            'properties': {'pixelSize': 400},
            'fields': 'pixelSize'
        }
    })
    
    # Header format
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 8},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': COLORS['header_dark'],
                    'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_gold']},
                    'horizontalAlignment': 'CENTER'
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
        }
    })
    
    for i, row in enumerate(rows[1:], start=1):
        if row[0] and 'WEEKLY TOTAL' not in row[0] and 'GRAND TOTAL' not in row[0]:  # Week header (Col A)
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': ws.id, 'startRowIndex': i, 'endRowIndex': i + 1, 'startColumnIndex': 0, 'endColumnIndex': 8},
                    'cell': {'userEnteredFormat': {'backgroundColor': COLORS['week_header_gold'], 'textFormat': {'bold': True}, 'horizontalAlignment': 'LEFT'}},
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            })
        elif 'WEEKLY TOTAL' in row[0]: 
            # Changed Alignment from CENTER to LEFT
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': ws.id, 'startRowIndex': i, 'endRowIndex': i + 1, 'startColumnIndex': 0, 'endColumnIndex': 8},
                    'cell': {'userEnteredFormat': {'backgroundColor': COLORS['total_gold'], 'textFormat': {'bold': True}, 'horizontalAlignment': 'LEFT'}}, 
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            })
        elif 'GRAND TOTAL' in row[0]:
            # Changed Alignment from CENTER to LEFT
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': ws.id, 'startRowIndex': i, 'endRowIndex': i + 1, 'startColumnIndex': 0, 'endColumnIndex': 8},
                    'cell': {'userEnteredFormat': {'backgroundColor': COLORS['grand_total_black'], 'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_gold']}, 'horizontalAlignment': 'LEFT'}},
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            })
            
    sh.batch_update({'requests': requests})
    print(f"  ✓ {tab_name} created")

def create_agents_tab(sh, data):
    """Create Agents Tab (All Time)"""
    tab_name = "Agents"
    print(f"\n👥 Creating {tab_name}...")
    
    # Aggregate by email (All Time)
    agent_data = {}
    
    if isinstance(data['account_analytics'], list):
        for day_stat in data['account_analytics']:
            email = day_stat.get('email_account')
            sent = day_stat.get('sent', 0)
            
            if email:
                if email not in agent_data:
                    agent_data[email] = {'total_sent': 0, 'active_days': 0}
                
                agent_data[email]['total_sent'] += sent
                if sent > 0:
                    agent_data[email]['active_days'] += 1
    
    for account in data['accounts']:
        email = account.get('email')
        if email and email not in agent_data:
            agent_data[email] = {'total_sent': 0, 'active_days': 0}
            
    sorted_agents = sorted(agent_data.items(), key=lambda x: x[1]['total_sent'], reverse=True)
    
    rows = []
    rows.append(['Agent Email', 'Total Emails Sent', 'Active Days', 'Avg/Day', 'Status'])
    
    total_sent = 0
    active_count = 0
    
    for email, stats in sorted_agents:
        total_sent += stats['total_sent']
        status = 'Active' if stats['total_sent'] > 0 else 'Inactive'
        if status == 'Active': active_count += 1
        
        avg_per_day = stats['total_sent'] / stats['active_days'] if stats['active_days'] > 0 else 0
        
        rows.append([
            email,
            f"{stats['total_sent']:,}",
            str(stats['active_days']),
            f"{avg_per_day:.0f}",
            status
        ])
        
    rows.append([f'TOTAL ({len(agent_data)} agents)', f"{total_sent:,}", '-', '-', f'{active_count} Active'])
    
    # Delete and recreate
    try:
        ws = sh.worksheet(tab_name)
        sh.del_worksheet(ws)
    except:
        pass
    try:
        ws_old = sh.worksheet("Agents 2026")
        sh.del_worksheet(ws_old)
    except:
        pass

    ws = sh.add_worksheet(title=tab_name, rows=200, cols=10)
    ws.update(values=rows, range_name='A1')
    
    # Formatting
    print("  → Applying formatting...")
    total_row_num = len(rows)
    requests = []
    
    # Header
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 5},
            'cell': {'userEnteredFormat': {'backgroundColor': COLORS['header_dark'], 'textFormat': {'bold': True, 'foregroundColor': COLORS['header_text_gold']}, 'horizontalAlignment': 'CENTER'}},
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
        }
    })
    
    # Alternating rows
    for i in range(1, total_row_num - 1):
        bg = COLORS['row_alt_gray'] if i % 2 == 0 else COLORS['bg_white']
        requests.append({
            'repeatCell': {
                'range': {'sheetId': ws.id, 'startRowIndex': i, 'endRowIndex': i + 1, 'startColumnIndex': 0, 'endColumnIndex': 5},
                'cell': {'userEnteredFormat': {'backgroundColor': bg}},
                'fields': 'userEnteredFormat(backgroundColor)'
            }
        })
        
    # Total Row
    requests.append({
        'repeatCell': {
            'range': {'sheetId': ws.id, 'startRowIndex': total_row_num - 1, 'endRowIndex': total_row_num, 'startColumnIndex': 0, 'endColumnIndex': 5},
            'cell': {'userEnteredFormat': {'backgroundColor': COLORS['total_gold'], 'textFormat': {'bold': True, 'foregroundColor': COLORS['text_black']}, 'horizontalAlignment': 'CENTER'}},
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
        }
    })
    
    sh.batch_update({'requests': requests})
    print(f"  ✓ {tab_name} created")

def main():
    print("=" * 80)
    print("CONNECT RESOURCES REPORT")
    print("=" * 80)
    
    api_key = os.getenv("INSTANTLY_API_KEY_CONNECT_RESOURCE")
    if not api_key:
        print("❌ Error: API key INSTANTLY_API_KEY_CONNECT_RESOURCE not found")
        return
    
    client = InstantlyClient(api_key)
    
    print("\n🔗 Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    try:
        sh = gc.open_by_key(SHEET_ID)
        print(f"  ✓ Connected to sheet: {sh.title}")
    except Exception as e:
        print(f"❌ Error connecting to sheet ID {SHEET_ID}: {e}")
        return
    
    # Fetch Data
    data = fetch_all_data(client)
    
    # Create Tabs
    create_master_dashboard(sh, data)
    create_campaigns_tab(sh, data)
    create_agents_tab(sh, data)
    
    # Re-order
    try:
        w1 = sh.worksheet("Master Dashboard")
        w2 = sh.worksheet("Campaigns 2026")
        w3 = sh.worksheet("Agents")
        sh.reorder_worksheets([w1, w2, w3])
        print("  ✓ Tabs reordered")
    except Exception as e:
        print(f"  ⚠️ Could not reorder tabs: {e}")
    
    print("\n✅ COMPLETE!")
    
    # --- Email Logic ---
    recipients_str = os.getenv("CONNECT_RESOURCES_REPORT_RECIPIENTS", "Jose@Luxvance.com")
    recipients = [r.strip() for r in recipients_str.split(",") if r.strip()]
    
    if recipients:
        # Calculate Date Range (Current Week)
        today = datetime.date.today()
        week_start = today - datetime.timedelta(days=today.weekday())
        week_end = week_start + datetime.timedelta(days=6)
        
        # Get ISO week number
        iso_year, iso_week, _ = week_start.isocalendar()
        
        # Metrics
        week_sent = 0
        week_replies = 0
        week_opps = 0
        
        all_time_sent = 0
        all_time_replies = 0
        all_time_opps = 0
        
        for c_id, days in data['campaign_analytics'].items():
            for day in days:
                d = datetime.datetime.strptime(day['date'], '%Y-%m-%d').date()
                
                # All-Time 
                all_time_sent += day.get('sent', 0)
                all_time_replies += day.get('unique_replies', 0)
                all_time_opps += day.get('opportunities', 0)
                
                # Current Week
                if week_start <= d <= week_end:
                    week_sent += day.get('sent', 0)
                    week_replies += day.get('unique_replies', 0)
                    week_opps += day.get('opportunities', 0)
        
        # Rates
        week_reply_rate = f"{(week_replies / week_sent * 100):.1f}%" if week_sent > 0 else "0.0%"
        week_opp_rate = f"{(week_opps / week_sent * 100):.1f}%" if week_sent > 0 else "0.0%"
        
        all_time_reply_rate = f"{(all_time_replies / all_time_sent * 100):.1f}%" if all_time_sent > 0 else "0.0%"
        all_time_opp_rate = f"{(all_time_opps / all_time_sent * 100):.1f}%" if all_time_sent > 0 else "0.0%"
        
        # Format Dates
        def suffix(d):
            return 'th' if 11<=d<=13 else {1:'st',2:'nd',3:'rd'}.get(d%10, 'th')
        
        start_fmt = f"{week_start.strftime('%A')} {week_start.day}{suffix(week_start.day)} of {week_start.strftime('%B')}"
        end_fmt = f"{week_end.strftime('%A')} {week_end.day}{suffix(week_end.day)} of {week_end.strftime('%B')} {week_end.year}"
        
        # HTML Content
        html_content = f"""
        <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
            <!-- Header -->
            <div style="background-color: #000000; padding: 20px; text-align: center;">
                <h2 style="color: #D4AF37; margin: 0; font-size: 24px;">Connect Resources Report - Week {iso_week}</h2>
                <div style="margin-top: 8px;">
                    <span style="color: #FFFFFF; font-size: 14px; background-color: #222; padding: 4px 12px; border-radius: 4px;">{start_fmt} to {end_fmt}</span>
                </div>
            </div>
            
            <!-- Weekly Metrics -->
            <div style="padding: 25px 20px 15px 20px;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <h3 style="margin: 0; color: #000; font-size: 16px; text-transform: uppercase;">Weekly Performance</h3>
                </div>
                <table style="width: 100%; text-align: center; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 10px; width: 33%;">
                            <div style="font-size: 28px; font-weight: bold; color: #000000;">{week_sent:,}</div>
                            <div style="font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Emails Sent</div>
                        </td>
                        <td style="padding: 10px; border-left: 1px solid #eee; border-right: 1px solid #eee; width: 33%;">
                            <div style="font-size: 28px; font-weight: bold; color: #D4AF37;">{week_replies:,}</div>
                            <div style="font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Replies</div>
                            <div style="font-size: 11px; color: #D4AF37; margin-top: 4px; font-weight: bold;">Rate: {week_reply_rate}</div>
                        </td>
                        <td style="padding: 10px; width: 33%;">
                            <div style="font-size: 28px; font-weight: bold; color: #000000;">{week_opps:,}</div>
                            <div style="font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Opportunities</div>
                            <div style="font-size: 11px; color: #000; margin-top: 4px; font-weight: bold;">Rate: {week_opp_rate}</div>
                        </td>
                    </tr>
                </table>
            </div>
            
            <!-- Divider -->
            <div style="height: 1px; background: linear-gradient(to right, #fff, #D4AF37, #fff); margin: 0 20px;"></div>
            
            <!-- All-Time Metrics -->
            <div style="padding: 20px 20px 30px 20px; background-color: #FAFAFA;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <h3 style="margin: 0; color: #666; font-size: 14px; text-transform: uppercase;">All-Time Performance</h3>
                </div>
                <table style="width: 100%; text-align: center; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 10px; width: 33%;">
                            <div style="font-size: 22px; font-weight: bold; color: #444;">{all_time_sent:,}</div>
                            <div style="font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 1px;">Total Sent</div>
                        </td>
                        <td style="padding: 10px; border-left: 1px solid #e0e0e0; border-right: 1px solid #e0e0e0; width: 33%;">
                            <div style="font-size: 22px; font-weight: bold; color: #BFAE58;">{all_time_replies:,}</div>
                            <div style="font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 1px;">Total Replies</div>
                            <div style="font-size: 10px; color: #BFAE58; margin-top: 2px;">Rate: {all_time_reply_rate}</div>
                        </td>
                        <td style="padding: 10px; width: 33%;">
                            <div style="font-size: 22px; font-weight: bold; color: #444;">{all_time_opps:,}</div>
                            <div style="font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 1px;">Total Opps</div>
                            <div style="font-size: 10px; color: #444; margin-top: 2px;">Rate: {all_time_opp_rate}</div>
                        </td>
                    </tr>
                </table>
                
                <div style="margin-top: 35px; text-align: center;">
                    <a href="https://docs.google.com/spreadsheets/d/{SHEET_ID}" style="background-color: #000000; color: #D4AF37; padding: 12px 30px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; border: 1px solid #D4AF37; font-size: 14px;">
                        View Master Dashboard
                    </a>
                </div>
            </div>
        </div>
        """
        
        send_email_report(
            subject=f"Connect Resources Report - Week {iso_week} ({week_start.strftime('%b %d')} to {week_end.strftime('%b %d')})",
            html_content=html_content,
            recipients=recipients
        )
    else:
        print("  ⚠️  No recipients found")

if __name__ == "__main__":
    main()
