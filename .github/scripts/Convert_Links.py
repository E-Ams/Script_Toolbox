import re
import os

def convert_wiki_links(content):
    """Convert wiki-style links to absolute GitHub wiki URLs"""
    
    # Pattern to match [text](link)
    pattern = r'\[([^\]]+)\]\(([^)]+)\)'
    
    def replace_link(match):
        text = match.group(1)
        link = match.group(2)
        
        print(f"Found link: '{text}' -> '{link}'")
        
        # Don't convert absolute URLs, anchor links, or mailto links
        if link.startswith(('http://', 'https://', '#', '/', 'mailto:')):
            print(f"  Skipping (absolute URL or anchor)")
            return match.group(0)
        
        # Convert wiki page links
        base_url = "https://github.com/E-Ams/Custom_Tools/wiki"
        new_link = f'[{text}]({base_url}/{link})'
        print(f"  Converting to: {new_link}")
        return new_link
    
    new_content = re.sub(pattern, replace_link, content)
    return new_content

def main():
    # Read the README.md file
    try:
        with open('README.md', 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print("README.md not found!")
        return
    
    print("=== Before conversion ===")
    print(content[:500] + "..." if len(content) > 500 else content)
    
    # Convert the links
    converted_content = convert_wiki_links(content)
    
    print("=== After conversion ===")
    print(converted_content[:500] + "..." if len(converted_content) > 500 else converted_content)
    
    # Write back to README.md
    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(converted_content)
    
    print("Successfully converted links in README.md")

if __name__ == '__main__':
    main()