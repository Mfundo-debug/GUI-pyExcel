# Data Collection and Visualization App

This program creates a GUI application using tkinter to collect user data and store it in an Excel file. The program also allows the user to load data from an existing Excel file, display the data in a treeview, and visualize the data using matplotlib.

## Further Explanation

This program creates a GUI application using tkinter to collect user data and store it in an Excel file. The program also allows the user to load data from an existing Excel file, display the data in a treeview, and visualize the data using matplotlib.

The program consists of two frames: a data entry form and a data display frame. The data entry form contains labels and entry fields for the user to input their name, age, gender, and hobbies. The form also includes a checkbox for the user to agree to the terms and conditions, and a submit button to add the data to the Excel file.

The program also includes a dark mode and light mode, which can be toggled using a menu. The program loads data from an existing Excel file using a load button, and displays the data in a treeview in the data display frame. The program also includes a visualize button to create a bar chart of the age distribution of the data using matplotlib.

The program is designed to be user-friendly and easy to use, with clear labels and error messages to guide the user through the data entry process.

## Features

- Data entry form with labels and entry fields for name, age, gender, and hobbies
- Checkbox for agreeing to terms and conditions
- Submit button to add data to Excel file
- Dark mode and light mode
- Load button to load data from existing Excel file
- Treeview to display loaded data
- Visualize button to create a bar chart of age distribution using matplotlib

## Requirements

- Python 3.x
- tkinter
- openpyxl
- matplotlib

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Mfundo-debug/GUI-pyExcel.git
```

2. Install the required packages:
```bash 
pip install -r requirements.txt
```

## Usage

1. Run the `pygui.py` file:
```bash
python py_gui.py
```

2. Fill out the data entry form and click the submit button to add data to the Excel file.
3. Click the load button to load data from an existing Excel file.
4. Click the visualize button to create a bar chart of the age distribution of the data.

## Credits

This program was created by [Mfundo Monchwe](https://github.com/Mfundo-debug).

## License

This program is licensed under the [MIT License](https://opensource.org/licenses/MIT).