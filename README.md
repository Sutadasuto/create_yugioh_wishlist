# Yu-Gi-Oh! Wishlist Creator

A Windows desktop app for building an Excel wishlist of cards from a Yu-Gi-Oh! TCG expansion. The app looks up expansion and card details from the official Yu-Gi-Oh! Card Database.

## Requirements

- Windows
- Python 3 installed, with the `py` launcher or Python available on `PATH`
- Tkinter included with the Python installation
- An internet connection

The launcher creates a project-local `.venv` and installs the packages from `requirements.txt` there. It does not install packages into your system Python.

## Run

Double-click **Create Wishlist.bat** and enter an expansion name when prompted. You can also pass the expansion name as an argument:

```bat
"Create Wishlist.bat" "Beyond the Brave"
```

The first run creates `.venv` and installs the required packages. Later runs reuse that environment.

## Use

The app loads the expansion's cards and lets you set the number of copies wanted for each detected rarity. Navigate through the cards, then choose **Finish and Export Excel** to create the wishlist.

Workbooks are saved in the project folder using a normalized expansion name, for example `beyond_the_brave.xlsx`. If a workbook for the expansion already exists, its saved quantities are loaded when the cards are matched by set card number or card name. Only quantities greater than zero are included when exporting.

## Dependencies

Python packages are listed in [`requirements.txt`](requirements.txt) and installed automatically by the batch launcher into `.venv`.