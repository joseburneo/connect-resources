"""
Helper to analyze inbox placement results by email provider
"""

def categorize_email_provider(email):
    """Categorize email by provider"""
    domain = email.split('@')[-1].lower()
    
    # Gmail/Google Workspace
    if 'gmail.com' in domain or 'googlemail.com' in domain:
        return 'Google'
    
    # Microsoft (Outlook, Hotmail, Live, Office365)
    if any(x in domain for x in ['outlook.', 'hotmail.', 'live.', 'msn.', 'office365.']):
        return 'Microsoft'
    
    # Yahoo
    if 'yahoo.' in domain or 'ymail.' in domain:
        return 'Yahoo'
    
    # Others
    return 'Others'


def analyze_provider_breakdown(test_results):
    """
    Analyze inbox placement by provider
    
    Args:
        test_results: Full test result from API with 'recipients' array
        
    Returns:
        dict with provider breakdown
    """
    
    if 'recipients' not in test_results:
        return None
    
    recipients = test_results['recipients']
    
    # Initialize counters
    provider_stats = {
        'Google': {'total': 0, 'inbox': 0, 'spam': 0, 'other': 0},
        'Microsoft': {'total': 0, 'inbox': 0, 'spam': 0, 'other': 0},
        'Yahoo': {'total': 0, 'inbox': 0, 'spam': 0, 'other': 0},
        'Others': {'total': 0, 'inbox': 0, 'spam': 0, 'other': 0}
    }
    
    # Process each recipient
    for recipient in recipients:
        if isinstance(recipient, str):
            # Simple string format - test not complete yet
            provider = categorize_email_provider(recipient)
            provider_stats[provider]['total'] += 1
        elif isinstance(recipient, dict):
            # Full result format with placement data
            email = recipient.get('email', '')
            placement = recipient.get('placement', 'other').lower()
            
            provider = categorize_email_provider(email)
            provider_stats[provider]['total'] += 1
            
            if 'inbox' in placement:
                provider_stats[provider]['inbox'] += 1
            elif 'spam' in placement:
                provider_stats[provider]['spam'] += 1
            else:
                provider_stats[provider]['other'] += 1
    
    # Calculate rates
    breakdown = {}
    for provider, stats in provider_stats.items():
        if stats['total'] > 0:
            breakdown[provider] = {
                'total': stats['total'],
                'inbox_rate': round((stats['inbox'] / stats['total']) * 100, 1),
                'spam_rate': round((stats['spam'] / stats['total']) * 100, 1),
                'other_rate': round((stats['other'] / stats['total']) * 100, 1),
                'inbox_count': stats['inbox'],
                'spam_count': stats['spam'],
                'other_count': stats['other']
            }
    
    return breakdown


def format_provider_breakdown_html(breakdown):
    """Format provider breakdown as HTML table"""
    
    if not breakdown:
        return "<p>Provider breakdown not available yet (test in progress)</p>"
    
    html = """
    <h3>📊 Provider Breakdown</h3>
    <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
        <tr style="background: #0b333d; color: white;">
            <th style="padding: 12px; text-align: left;">Provider</th>
            <th style="padding: 12px; text-align: center;">Total</th>
            <th style="padding: 12px; text-align: center;">📥 Inbox</th>
            <th style="padding: 12px; text-align: center;">🚫 Spam</th>
            <th style="padding: 12px; text-align: center;">📂 Other</th>
        </tr>
    """
    
    def get_color(rate):
        if rate >= 85:
            return "#28a745"
        elif rate >= 75:
            return "#ffc107"
        else:
            return "#dc3545"
    
    for provider in ['Google', 'Microsoft', 'Yahoo', 'Others']:
        if provider in breakdown:
            stats = breakdown[provider]
            html += f"""
        <tr style="background: #f8f9fa;">
            <td style="padding: 12px;"><strong>{provider}</strong></td>
            <td style="padding: 12px; text-align: center;">{stats['total']}</td>
            <td style="padding: 12px; text-align: center; background: {get_color(stats['inbox_rate'])}; color: white;">
                <strong>{stats['inbox_rate']}%</strong><br>
                <small>({stats['inbox_count']})</small>
            </td>
            <td style="padding: 12px; text-align: center;">
                {stats['spam_rate']}%<br>
                <small>({stats['spam_count']})</small>
            </td>
            <td style="padding: 12px; text-align: center;">
                {stats['other_rate']}%<br>
                <small>({stats['other_count']})</small>
            </td>
        </tr>
            """
    
    html += """
    </table>
    """
    
    return html


def format_provider_breakdown_text(breakdown):
    """Format provider breakdown as plain text"""
    
    if not breakdown:
        return "Provider breakdown not available yet (test in progress)"
    
    text = "\n📊 Provider Breakdown:\n\n"
    
    for provider in ['Google', 'Microsoft', 'Yahoo', 'Others']:
        if provider in breakdown:
            stats = breakdown[provider]
            text += f"{provider:12} | Total: {stats['total']:3} | "
            text += f"Inbox: {stats['inbox_rate']:5.1f}% ({stats['inbox_count']:2}) | "
            text += f"Spam: {stats['spam_rate']:5.1f}% ({stats['spam_count']:2}) | "
            text += f"Other: {stats['other_rate']:5.1f}% ({stats['other_count']:2})\n"
    
    return text
