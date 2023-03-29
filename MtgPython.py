import requests
from bs4 import BeautifulSoup
import win32com.client as win32
import os
import re

# Start Microsoft Publisher
publisher = win32.gencache.EnsureDispatch('Publisher.Application')

# Create a new publication
publication = publisher.Documents.Add()

# Define the URL to search
url = 'https://scryfall.com/search?q='

# Define the file containing the search terms
file_name = 'search_terms.txt'

# Open the file containing the search terms
with open(file_name, 'r', encoding='utf-8') as f:
    # Read the lines and filter out lines that don't start with a number followed by 'x'
    search_terms = [line.strip() for line in f if re.search(r'\d+x', line)]
    #print(search_terms)

# Define the size of each card
card_width = 2.5 * 72
card_height = 3.483 * 72

# Define the top and left margin of the card grid
top_margin = 0.25 * 72
left_margin = 0.5 * 72

# Initialize the cards on page to 0
cards_on_page = 0

# Loop through each search term
for search_term in search_terms:
    # Split the line into the number of times and the search term
    count, search_term = search_term.split('x', maxsplit=1)
    count = int(count)

    # Replace any spaces in the search term with '+'
    search_term = '"' + search_term.strip().replace(' ', '+') + '"'

    for i in range(count):
        # Combine the URL and search term
        full_url = url + '+!' + search_term

        # Send a request to the URL and get the response
        response = requests.get(full_url)

        # Parse the HTML using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find the image URL in the HTML
        image_url_element = soup.find('meta', {'property': 'og:image'})
        if image_url_element:
            image_url = image_url_element['content']
            # Download the image
            response = requests.get(image_url)
            # Save the image to a file
            with open('image.jpg', 'wb') as f:
                f.write(response.content)
            # Get the full path to the image file
            image_path = os.path.abspath('image.jpg')
            # Add the image to the publication
            shape = publication.Pages(publication.Pages.Count).Shapes.AddPicture(
                Filename=image_path,  # Specify the filename
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left_margin + (cards_on_page % 3) * card_width,
                Top=top_margin + (cards_on_page // 3) * card_height,
                Width=card_width,
                Height=card_height
            )
            # Replace the picture with the downloaded image
            shape.PictureFormat.Replace(image_path)
            # Increment the number of cards on the current page
            cards_on_page += 1
        else:
            print(f"No image URL found for search term: {search_term}")

        # If we have added 9 cards, create a new page
        if cards_on_page == 9:
            publication.Pages.Add(Count=1, After=publication.Pages.Count)
            cards_on_page = 0

# Print out finished message
print("Finished exporting images")
# Save the publication
#publication.SaveAs('Magic Cards.pub')
