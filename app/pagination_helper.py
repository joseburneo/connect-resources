"""
Helper function to fetch ALL items from paginated Instantly API endpoints
"""

def fetch_all_paginated(client, endpoint, initial_params=None, max_pages=20):
    """
    Fetch all items from a paginated Instantly API endpoint.
    
    Args:
        client: InstantlyClient instance
        endpoint: API endpoint path (e.g., '/accounts', '/campaigns')
        initial_params: Dict of initial query parameters (optional)
        max_pages: Maximum number of pages to fetch (safety limit)
    
    Returns:
        list: All items from all pages
    """
    all_items = []
    cursor = None
    page = 1
    
    while page <= max_pages:
        # Build params
        params = initial_params.copy() if initial_params else {}
        if cursor:
            params['starting_after'] = cursor
        
        # Fetch page
        response = client.request("GET", endpoint, params=params)
        items = response.get('items', [])
        
        if not items:
            break
        
        all_items.extend(items)
        print(f"  📄 Page {page}: {len(items)} items (total so far: {len(all_items)})")
        
        # Check for next page
        cursor = response.get('next_starting_after')
        if not cursor:
            break
        
        page += 1
    
    return all_items
