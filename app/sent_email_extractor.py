"""
Helper to extract real sent email copy from campaigns
Uses actual sent emails with rendered variables instead of templates
"""

def get_sent_email_from_campaign(client, campaign_id, limit=20):
    """
    Fetch a real sent email from a campaign
    
    Args:
        client: InstantlyClient instance
        campaign_id: Campaign ID to fetch emails from
        limit: Number of emails to check (default 20)
        
    Returns:
        dict with subject and body, or None if no emails found
    """
    try:
        # Get sent emails from campaign
        # ue_type: 1 = sent from campaign, 2 = reply
        response = client.request('GET', '/emails', params={
            'limit': limit,
            'campaign_id': campaign_id
        })
        
        emails = response.get('items', [])
        
        # Filter for sent emails only (not replies)
        sent_emails = [e for e in emails if e.get('ue_type') == 1]
        
        if not sent_emails:
            return None
        
        # Use the first sent email (most recent)
        email = sent_emails[0]
        
        # Extract subject
        subject = email.get('subject', '')
        
        # Extract body
        body_data = email.get('body', {})
        body = ''
        
        if isinstance(body_data, dict):
            # Prefer HTML body, fallback to text
            body = body_data.get('html', body_data.get('text', ''))
        elif isinstance(body_data, str):
            body = body_data
        
        return {
            'subject': subject,
            'body': body,
            'to': email.get('to_address_email_list', ''),
            'timestamp': email.get('timestamp_email', ''),
            'is_rendered': True  # Flag to indicate this is real data
        }
        
    except Exception as e:
        print(f"⚠️  Error fetching sent emails: {e}")
        return None


def get_best_campaign_copy_rendered(client):
    """
    Fetch REAL sent email copy from best performing campaign
    Uses actual sent emails with rendered variables (no {{firstName}})
    
    Args:
        client: InstantlyClient instance
        
    Returns:
        dict with subject, body, and campaign name
    """
    from app.pagination_helper import fetch_all_paginated
    
    SAFE_TEMPLATE = {
        "subject": "Quick question",
        "body": """Hi there,

I wanted to reach out about a potential opportunity.

Would you be open to a brief conversation?

Best regards"""
    }
    
    try:
        # Get all campaigns
        all_campaigns = fetch_all_paginated(client, "/campaigns")
        active = [c for c in all_campaigns if c.get('status') == 1]
        
        if not active:
            return {**SAFE_TEMPLATE, "name": "Safe Template (No Active Campaigns)"}
        
        # Sort by reply rate (best metric)
        best = max(active, key=lambda c: c.get('reply_rate', c.get('open_rate', 0)))
        campaign_id = best['id']
        campaign_name = best.get('name', 'Best Campaign')
        
        print(f"📧 Fetching real sent email from: {campaign_name}")
        
        # Get a real sent email from this campaign
        sent_email = get_sent_email_from_campaign(client, campaign_id)
        
        if sent_email:
            print(f"✅ Found real sent email!")
            print(f"   Subject: {sent_email['subject'][:60]}...")
            print(f"   To: {sent_email['to']}")
            print(f"   Body length: {len(sent_email['body'])} chars")
            print(f"   ✨ Variables are RENDERED (no {{firstName}}, actual names!)")
            
            return {
                "subject": sent_email['subject'],
                "body": sent_email['body'],
                "name": campaign_name,
                "is_rendered": True
            }
        else:
            print(f"⚠️  No sent emails found, using template fallback")
            
            # Fallback to template extraction
            sequences = best.get('sequences', [])
            if sequences and len(sequences) > 0:
                steps = sequences[0].get('steps', [])
                if steps and len(steps) > 0:
                    variants = steps[0].get('variants', [])
                    if variants and len(variants) > 0:
                        variant = variants[0]
                        return {
                            "subject": variant.get('subject', SAFE_TEMPLATE['subject']) or SAFE_TEMPLATE['subject'],
                            "body": variant.get('body', SAFE_TEMPLATE['body']) or SAFE_TEMPLATE['body'],
                            "name": campaign_name,
                            "is_rendered": False
                        }
            
            return {**SAFE_TEMPLATE, "name": campaign_name}
            
    except Exception as e:
        print(f"⚠️  Error: {e}")
        return {**SAFE_TEMPLATE, "name": "Safe Template (Error)"}
