# Yagooglesearch Enhanced Google Search Scraper

A Python script with a Tkinter-based GUI that utilizes `yagooglesearch` for efficient Google searches via the Custom Search JSON API. The script allows you to control search parameters, manage sleep intervals, and organizes results effortlessly in an Excel file.

## Prerequisites

Before using the script, ensure you have the following:

- Python installed (version 3.6 or higher)
- Required Python libraries installed (`openpyxl`, `tkinter`, `random`, `yagooglesearch`)

You can install the required libraries using:

```bash
pip install openpyxl tkinter yagooglesearch
```

## Getting Started

1. Clone the repository to your local machine:

```bash
git clone https://github.com/your-username/yagooglesearch-enhanced-google-search-scraper.git
```

2. Navigate to the project directory:

```bash
cd yagooglesearch-enhanced-google-search-scraper
```

3. Run the script:

```bash
python yagooglesearch_enhanced_google_search_scraper.py
```

## Usage

1. The GUI prompts you to select a keywords file, which should contain one keyword per line.
2. Choose the output folder where the Excel file with the search results will be saved.
3. Click the "Run" button to initiate the Google searches.

The script generates an Excel file with the search results, including the keyword, rank, title, URL, and description.

## Advanced Configuration

For advanced users, you can further customize the script by modifying parameters in the script or exploring the `yagooglesearch` library.

## Important Note

- Keep your API key and Search Engine ID confidential. Do not expose them on public repositories or share them with unauthorized users.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Special thanks to the creators of the Python libraries (`yagooglesearch`) used in this script.