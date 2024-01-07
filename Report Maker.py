import tkinter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pyproj
import simplekml
import os
from shapely.geometry import Polygon
import matplotlib.pyplot as plt
import IPython
from matplotlib.ticker import ScalarFormatter
from geopy.distance import geodesic
from math import atan2, degrees, sin, cos, radians
import docx
from docx import Document
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
import base64
from io import BytesIO


# BACKGROUND_COLOR = "#2F4F4F"
BACKGROUND_COLOR = "#2F4F5F"
# BACKGROUND_COLOR = "#527bb0"
FOREGROUND_COLOR = "#FFFFFF"
FONT = "Ebrima 14"
FONT_ONLY = "Ebrima"
BUTTON_CLICK_COLOR = "#98F5FF"


root = Tk()
root.title("Report Maker")
root.config(bg=BACKGROUND_COLOR)
root.geometry("750x600")
root.resizable(False,False)

label_font = ("Helvetica", 12)
entry_font = ("Helvetica", 12)

# Set background color for labels
label_bg_color = "#4285F4"

prompt_label = Label(root, text="•Fill in the form and upload an excel csv file with the coordinaties\n"
                               "•The Coordinates of the excel csv should be arranged in Northings Eastings Label\n"
                                "•The coordinate system should be in Ghana War Office for plotting on google earth", font="Ebrima 14",
                                bg=BACKGROUND_COLOR, fg=FOREGROUND_COLOR, justify="left")
prompt_label.grid(column=0, row=0, columnspan=2, pady=20)

Label(root, text="Client Name:", font=label_font, bg=label_bg_color, fg="white").grid(row=1, column=0,
                                                                                              sticky="ew", pady=5,
                                                                                              padx=5)
client_name_entry = Entry(root, font=entry_font)
client_name_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=5)

Label(root, text="Nationality:", font=label_font, bg=label_bg_color, fg="white").grid(row=2, column=0,
                                                                                              sticky="ew", pady=5,
                                                                                              padx=5)
nationality_entry = Entry(root, font=entry_font)
nationality_entry.grid(row=2, column=1, sticky="ew", pady=5, padx=5)

Label(root, text="District:", font=label_font, bg=label_bg_color, fg="white").grid(row=3, column=0, sticky="ew",
                                                                                           pady=5, padx=5)
district_entry = Entry(root, font=entry_font)
district_entry.grid(row=3, column=1, sticky="ew", pady=5, padx=5)

Label(root, text="Region:", font=label_font, bg=label_bg_color, fg="white").grid(row=4, column=0, sticky="ew",
                                                                                         pady=5, padx=5)
region_entry = Entry(root, font=entry_font)
region_entry.grid(row=4, column=1, sticky="ew", pady=5, padx=5)

Label(root, text="Regional Number:", font=label_font, bg=label_bg_color, fg="white").grid(row=5, column=0,
                                                                                                  sticky="ew", pady=5,
                                                                                                  padx=5)
regional_number_entry = Entry(root, font=entry_font)
regional_number_entry.grid(row=5, column=1, sticky="ew", pady=5, padx=5)


Label(root, text="Date:", font=label_font, bg=label_bg_color, fg="white").grid(row=6, column=0,
                                                                                                  sticky="ew", pady=5,
                                                                                                  padx=5)
date_entry = Entry(root, font=entry_font)
date_entry.grid(row=6, column=1, sticky="ew", pady=5, padx=5)



Label(root, text="Surveyor Name:", font=label_font, bg=label_bg_color, fg="white").grid(row=7, column=0,
                                                                                                  sticky="ew", pady=5,
                                                                                                  padx=5)
surveyor_entry = Entry(root, font=entry_font)
surveyor_entry.grid(row=7, column=1, sticky="ew", pady=5, padx=5)

excel_entry = Entry(root,font=entry_font, width=50, bg=BACKGROUND_COLOR)
excel_entry.grid(row=11, column=0, columnspan=2)

def browse_excel():
    excel_entry.delete(0, END)
    filename = filedialog.askopenfilename(initialdir="/", title="Select Co-ordinate File",
                                          filetypes=(("csv file", "*.csv"),
                                                     ))
    excel_entry.insert(END, string=filename)
   # file_directory = excel_entry.get()
    # file_d_label = Label(root, text=file_directory, font=label_font, bg=label_bg_color)
    # excel_entry.grid(row=11, column=0, columnspan=2)
    print(excel_entry.get())

browse_button = Button(root, text="Browse", command=browse_excel, width=50, activebackground=BUTTON_CLICK_COLOR)
browse_button.grid(row=8, column=0, columnspan=2, pady=20)

from shapely.geometry import Polygon
import matplotlib.pyplot as plt


def plot_coordinates(coordinates):
    # Extracting coordinates from the input list (excluding labels)
    points = [(float(coord.split(',')[1]), float(coord.split(',')[0])) for coord in coordinates]

    # Extracting labels from the third object in the list
    labels = [coord.split(',')[2].strip() for coord in coordinates]

    # Creating a Polygon from the points
    polygon = Polygon(points)

    # Extracting x and y coordinates for plotting
    x, y = polygon.exterior.xy

    # Plotting the map-style plot
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.plot(x, y, color='blue', alpha=0.7, linewidth=2, solid_capstyle='round', zorder=2)
    ax.fill(x, y, color='lightblue', alpha=0.4, edgecolor='blue', linewidth=2, zorder=1)

    # Marking the vertices with labels from the third object in the list
    for point, label in zip(points, labels):
        ax.scatter(point[0], point[1], color='red', marker='o', label='Vertices', zorder=3)
        ax.text(point[0], point[1], label, fontsize=8, ha='right', va='bottom', color='black')

    # Designing the plot to look like a map
    ax.set_title('Map of Survey Area (Close This plot to view the report)')
    ax.set_xlabel('Easting')
    ax.set_ylabel('Northing')

    ax.legend()

    # Show plot
    plt.show()



#
# # obtaining information from entry boxes
# client_name = client_name_entry.get()
# client_nationality = nationality_entry.get()
# client_locality = district_entry.get()
# client_region = region_entry.get()
# client_regional_number = regional_number_entry.get()
# client_work_date = date_entry.get()
# client_surveyor = surveyor_entry.get()


# Writing Report


def write_report(input_area, input_perimeter, input_coordinates):
    print("THS IS  A FISH")
    global client_name, client_nationality, client_locality, client_region, client_regional_number, client_work_date, client_surveyor
    with open('report.txt', 'r', encoding="utf-8") as file_report:
        content = file_report.read()

    # obtaining information from entry boxes
    client_name = client_name_entry.get()
    client_nationality = nationality_entry.get()
    client_locality = district_entry.get()
    client_region = region_entry.get()
    client_regional_number = regional_number_entry.get()
    client_work_date = date_entry.get()
    client_surveyor = surveyor_entry.get()

    content = content.replace("WORK_DATE", client_work_date)
    content = content.replace("WORK_SURVEYOR", client_surveyor)
    content = content.replace("CLIENT_NAME", client_name)
    content = content.replace("CLIENT_NATIONALITY", client_nationality)
    content = content.replace("CLIENT_LOCALITY", client_locality)
    content = content.replace("CLIENT_REGION", client_region)
    content = content.replace("CLIENT_PERSONAL_NUMBER", client_regional_number)
    content = content.replace("WORK_DATE", client_work_date)
    content = content.replace("WORK_DATE", client_work_date)
    content = content.replace("LAND_AREA", str(input_area))
    content = content.replace("LAND_PERIMETER", str(input_perimeter))
    content = content.replace("LAND_COORDINATES", str(input_coordinates))
    print(content)

    with open("final_report.txt", mode="w") as final:
        final.write(content)

    #writing to the word file
    with open("final_report.txt", mode="r") as final_read:
        word_content = final_read.read()

    doc = Document()
    doc.add_paragraph(word_content)

    doc.save("final_report.docx")

def open_word_file():
    word_file_path = 'final_report.docx'
    os.system(f'start {word_file_path}')


def generate_report():
    try:
        print(f"this is the problem {excel_entry.get()}")
        csv_file = excel_entry.get()
        with open(csv_file) as csv:
            cord_line = csv.readlines()
            print(f" this is cord line {cord_line}")
            points = [(float(coord.split(',')[0]), float(coord.split(',')[1])) for coord in cord_line]

            # creating a polygon from the points
            polygon = Polygon(points)

            # calculating area and perimeter
            area = round(polygon.area/43560,2)
            perimeter = round(polygon.length,2)

            print(f"this is area {area} acres and this is perimeter: {perimeter} feets")

        for co_ordinates in cord_line:
            only_co_ordinates = co_ordinates.strip("\n")
            co_ordinates_list = only_co_ordinates.split(",")
            print(co_ordinates_list)

        plot_coordinates(cord_line)
        write_report(input_area=area, input_perimeter=perimeter,input_coordinates=points)

        open_word_file()

    except FileNotFoundError as e:
        print("eRROR")
        messagebox.showerror(title="Error",message=e)
        # print(e)

report_button = Button(root, text="Generate Report", command=generate_report, width=50, activebackground=BUTTON_CLICK_COLOR)
report_button.grid(row=9, column=0, pady=20, columnspan=2)


# Plotting on google earth

# Plotting the plot
def plot():
    is_success = True
    csv_file = excel_entry.get()

    output_file = filedialog.asksaveasfile(title="Select where you want to save the file",
                                           confirmoverwrite=True)
    output_data = output_file.name
    print(output_data)

    with open(csv_file) as csv:
        cord_line = csv.readlines()

    co_ordinate_pairs = []
    with open(f"{output_data}.txt", mode="a") as google_coordinates:
        google_coordinates.write(f"Longitude,Latitude,Label\n")
        for co_ordinates in cord_line:
            only_co_ordinates = co_ordinates.strip("\n")
            co_ordinates_list = only_co_ordinates.split(",")
            print(co_ordinates_list)
            try:
                source_eastings = float(co_ordinates_list[1])
                source_northings = float(co_ordinates_list[0])
                war_to_wgs_84 = pyproj.Transformer.from_crs(2136, 4326)
                wgs_value = war_to_wgs_84.transform(source_eastings, source_northings)
                # putting converted wgs coordinate in a tuple along with the names
                converted_coord = (wgs_value[1], wgs_value[0], co_ordinates_list[2])
                decimal_degree = f"{round(wgs_value[1], 6)}, {round(wgs_value[0], 6)}, {co_ordinates_list[2]}"
                google_coordinates.write(f"{decimal_degree}\n")
                co_ordinate_pairs.append(converted_coord)
                # messagebox.showinfo(title="Success", message="Your co-ordinates have been successfully converted and is"
                #                                              " being plotted on google earth. Please make sure you have"
                #                                              " google earth installed on your pc")
                is_success = True
            except ValueError:
                is_success = False
                pass

    print(co_ordinate_pairs)

    # plotting on to google earth
    map_kml = simplekml.Kml()
    # plotting points

    line_points = []
    for co_ordinate in co_ordinate_pairs:
        map_kml.newpoint(name=co_ordinate[2], coords=[(co_ordinate[0], co_ordinate[1])])
        line_co_ordinate = (co_ordinate[0], co_ordinate[1])
        line_points.append(line_co_ordinate)

    # plotting lines
    pol = map_kml.newpolygon(name="boundary", outerboundaryis=line_points)
    pol.style.polystyle.color = '990000ff'

    map_kml.save(f"{output_data}.kml")

    os.startfile(os.path.abspath(f"{output_data}.kml"))

    if is_success:
        messagebox.showinfo(title="Success", message="your coordinates have been successfully plotted")
    else:
        messagebox.showerror(title="Error", message="There was an error. Please make sure you have google earth installed and also"
                                                    ". Make sure there are no numericals in the label column."
                                                    )

plot_button = Button(root, text="Plot on Google Earth", command=plot, width=50, activebackground=BUTTON_CLICK_COLOR)
plot_button.grid(row=10, column=0, pady=20, columnspan=2)



mainloop()
