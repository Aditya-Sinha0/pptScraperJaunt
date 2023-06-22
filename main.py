import requests
import csv
import os
import langid

# Google Custom Search API key
api_key = os.environ.get('API_key')

# Check if the API key is available
if api_key is None:
    raise ValueError('API_KEY environment variable is not set.')

# Google Custom Search Engine ID
search_engine_id = 'c3f38d667dba34e3f'

def get_language_name(code):
    # Map language codes to full language names
    language_names = {
        'af': 'Afrikaans',
        'ar': 'Arabic',
        'bg': 'Bulgarian',
        'bn': 'Bengali',
        'ca': 'Catalan',
        'cs': 'Czech',
        'cy': 'Welsh',
        'da': 'Danish',
        'de': 'German',
        'el': 'Greek',
        'en': 'English',
        'es': 'Spanish',
        'et': 'Estonian',
        'fa': 'Persian',
        'fi': 'Finnish',
        'fr': 'French',
        'gu': 'Gujarati',
        'he': 'Hebrew',
        'hi': 'Hindi',
        'hr': 'Croatian',
        'hu': 'Hungarian',
        'id': 'Indonesian',
        'it': 'Italian',
        'ja': 'Japanese',
        'kn': 'Kannada',
        'ko': 'Korean',
        'lt': 'Lithuanian',
        'lv': 'Latvian',
        'mk': 'Macedonian',
        'ml': 'Malayalam',
        'mr': 'Marathi',
        'ne': 'Nepali',
        'nl': 'Dutch',
        'no': 'Norwegian',
        'pa': 'Punjabi',
        'pl': 'Polish',
        'pt': 'Portuguese',
        'ro': 'Romanian',
        'ru': 'Russian',
        'sk': 'Slovak',
        'sl': 'Slovenian',
        'so': 'Somali',
        'sq': 'Albanian',
        'sv': 'Swedish',
        'sw': 'Swahili',
        'ta': 'Tamil',
        'te': 'Telugu',
        'th': 'Thai',
        'tl': 'Tagalog',
        'tr': 'Turkish',
        'uk': 'Ukrainian',
        'ur': 'Urdu',
        'vi': 'Vietnamese',
        'zh-cn': 'Chinese (Simplified)',
        'zh-tw': 'Chinese (Traditional)'
    }

    return language_names.get(code, 'Unknown language')


def search_ppt_files(query, num_results=10):
    results_per_request = 10  # Number of results per API request

    # Calculate the number of API requests needed based on the desired number of results
    num_requests = (num_results - 1) // results_per_request + 1

    # Create a CSV file to store the results
    with open('ppt_results.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['URL', 'Language'])  # Write the header row to the CSV file

        # Perform the API requests
        for request_num in range(num_requests):
            start_index = request_num * results_per_request + 1

            # Search for .ppt files based on query using Google Custom Search JSON API
            url = f"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={search_engine_id}&q={query}&fileType=ppt&start={start_index}"
            response = requests.get(url)
            results = response.json()

            # Iterate over the search results and retrieve the .ppt files
            if 'items' in results:
                for item in results['items']:
                    file_url = item['link']
                    file_id = file_url.split('/')[-2]

                    # Detect the language of the snippet
                    language_code, _ = langid.classify(item['snippet'])
                    language = get_language_name(language_code)

                    # Write the URL and language to the CSV file
                    writer.writerow([file_url, language])


# Provide the query keyword and number of results you want
query_keyword = 'トマト'
num_results = 20
search_ppt_files(query_keyword, num_results)