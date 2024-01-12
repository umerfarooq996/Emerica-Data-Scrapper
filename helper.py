import re

def remove_double_spaces(text):
    return re.sub(r' +', ' ', text)

def singularize(word):
    irregular_mappings = {
        'men': 'man',
        'women': 'woman',
        # Add more irregular mappings as needed
    }

    # Check if the word exists in irregular mappings
    if word.lower() in irregular_mappings:
        return irregular_mappings[word.lower()]

    # Define some common suffixes that indicate plural form
    common_suffixes = ['s', 'es']

    # Check if the word ends with any common suffixes
    for suffix in common_suffixes:
        if word.lower().endswith(suffix):
            # Remove the suffix to get the singular form
            return word[:-len(suffix)]

    # If no special case or common suffix is found, return the same word
    return word


def extract_style_code(images, sku):
    def get_code(link):
        try:
            pattern = f"{sku}-(\w+)-"
            match = re.search(pattern,link)
            if match:
                return match.group(1)
            else:
                return
        except Exception as e:
            print(f"Error: {e}")
        
    for i in images:
        code=get_code(i)
        if code:
            return code

def switch_words(input_str):
    # Split the input string into words
    words = input_str.split()

    # Check if the input string has at least three words
    if len(words) >= 3:
        # Find the position of the word "Boys"
        boys_index = words.index("Boys")

        # Check if "Boys" is not the last word
        if boys_index < len(words) - 1:
            # Switch the positions of "Boys" and the second word
            words[1], words[boys_index] = words[boys_index], words[1]

            # Join the words back together to form the new string
            result_str = " ".join(words)
            return result_str

    # If there are less than three words or "Boys" is not found, return the original string
    return input_str

def getPrice(p):
    if p:
        p = p.replace('$', '').strip()
        p = round(float(p))
        p = int(p)
        return p