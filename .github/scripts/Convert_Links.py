import re
import os

def convert_wiki_links(content):
    """Convert wiki-style links to absolute GitHub wiki URLs"""
    
    # More flexible pattern that handles various markdown link formats
    patterns = [
        r'\[([^\]]+)\]\(([^)#]+)\)',  # Standard [text](link)
        r'\[([^\]]+)\]\[([^\]]+)\]',  # Reference style [text][ref]
    ]
    
    converted_content = content
    
    # Handle standard links
    def replace_standard_link(match):
        text = match.group(1)
        link = match.group(2).strip()
        
        print(f"Standard link: '{text}' -> '{link}'")
        
        # Skip if already absolute or special link
        if any(link.startswith(prefix) for prefix in 
               ['http://', 'https://', '#', '/', 'mailto:']):
            return match.group(0)
        
        # Convert to absolute wiki URL
        base_url = "https://github.com/E-Ams/Custom_Tools/wiki"
        return f'[{text}]({base_url}/{link})'
    
    converted_content = re.sub(patterns[0], replace_standard_link, converted_content)
    
    return converted_content

def main():
    # Read the README.md file
    try:
        with open('README.md', 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print("README.md not found!")
        return
    
    # Convert the links
    converted_content = convert_wiki_links(content)
    
    # Write back to README.md
    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(converted_content)
    
    print("Link conversion completed")

if __name__ == '__main__':
    main()