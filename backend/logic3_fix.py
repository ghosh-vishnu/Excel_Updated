# New Logic 3 implementation for proper nested list handling

def new_logic3_implementation(para, text, is_bold, html_output, list_levels, inside_list):
    """New Logic 3 implementation with proper nested list handling"""
    
    # Check if this is a list item
    is_word_list_item = is_list_item(para)
    has_bullet_chars = any(char in text for char in ['•', '-', '–', '○', '◦', '‣', '▪', '▫', '*', '+'])
    has_numbering = re.match(r'^\d+[\.\)]', text)
    is_list_item_detected = is_word_list_item or has_bullet_chars or has_numbering
    
    if is_list_item_detected:
        # ANY list item (bold or non-bold) - handle nested structure
        # Get plain text without bold formatting for list items
        formatted_content = para.text.strip()
        if formatted_content:
            # Get list level for nested lists
            current_level = get_list_level(para)
            print(f"DEBUG: current_level={current_level}, list_levels={list_levels}, inside_list={inside_list}")
            
            # Close deeper levels if we're going back to a higher level
            while len(list_levels) > current_level:
                html_output.append("</ul>")
                list_levels.pop()
                print(f"DEBUG: Closed deeper level, list_levels={list_levels}")
            
            # Open new level if needed
            while len(list_levels) < current_level:
                html_output.append("<ul>")
                list_levels.append(len(list_levels))
                print(f"DEBUG: Opened new level, list_levels={list_levels}")
            
            # Ensure we have at least one list open
            if not list_levels:
                html_output.append("<ul>")
                list_levels.append(0)
                inside_list = True
                print(f"DEBUG: Started first list, list_levels={list_levels}")
            
            html_output.append(f"<li><p>{formatted_content}</p></li>")
            print(f"DEBUG LOGIC 3: Added <li><p> for list item (bold={is_bold}, level={current_level}): {formatted_content[:30]}...")
            
    elif is_bold and not is_list_item_detected:
        # Bold heading (NOT in list) - close all open lists first
        while list_levels:
            html_output.append("</ul>")
            list_levels.pop()
        inside_list = False
        
        heading_text = clean_heading(text)
        if heading_text:
            html_output.append(f"\n<strong>{heading_text}</strong>")
            print(f"DEBUG LOGIC 3: Added <strong> for bold heading: {heading_text[:30]}...")
            
    else:
        # Regular paragraph text - close all open lists
        while list_levels:
            html_output.append("</ul>")
            list_levels.pop()
        inside_list = False
        
        formatted_content = runs_to_html_with_links(para.runs)
        if formatted_content:
            html_output.append(f"<p>{formatted_content}</p>")
            print(f"DEBUG LOGIC 3: Added <p> for regular text: {formatted_content[:30]}...")
    
    return html_output, list_levels, inside_list
