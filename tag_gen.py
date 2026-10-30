from openpyxl import Workbook
from tkinter import *
from tkinter import ttk
import datetime
from pathlib import Path

#INFO REQUIRED
    #is brand a dropdown?
        #maybe top10 dropdown and an "other" option and field
    #consider confirm and submit split
        #confirm shows a preview row, submits builds and deploys sheet

#TODO
    #Clear All button
    #Feature add to list (and remove)
        #functional, add better delete system
    #Pant size list nR ?
        #nR(egular), nS(hort), nT(all)
        #probably not worth adding, seems uncommon
        #custom size box?
    #Womens (shoe) size toggle
    #Multiple Colours
    #Multiple Workwear Categories
    
   

#creates and populates workbook
def submit():
    #create workbook object
    wb = Workbook()

    #grab the active worksheet
    ws = wb.active
    
    type = selected_type_category.get()
    #This header is different if product is shoes
    if type == "Shoes": 
        type_cat = "Footwear Type"
    else:
       type_cat = "Workwear Categories"
    
    #header line, required for Nomad formatting
    ws.append(["Product ID", "Brand", "Colour", type_cat, "Features", "Size"])
    
    #get values from user input
    code_value = input_box_code.get()
    brand_value = input_box_brand.get()
    colour_value = selected_colour.get()
    type_value = input_box_type.get()
    
    #features needs to be one string, Nomad formatting uses | as a seperator
    features_value = ""
    for f in features:
        features_value += f +" | "
    features_value = features_value[0:-3]
    
    #new row for each size, size is appended in its own column as well as on ID
    for size in selected_sizes:
        size_disp = size
        if type == "Pants":
            size_disp = size.replace("/", "W x ") + "L"
        ws.append([code_value+"-"+size, brand_value, colour_value, type_value, features_value, size_disp])
    
    # Make Outputs folder if does not exist
    dir_path = Path("./outputs")
    dir_path.mkdir(parents=True, exist_ok=True)
    
    #name worksheet, required for Nomand import
    ws.title = "Tags"

    # Save the file
    # Generates filename using Product ID and current time, for unique filenames
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"./outputs/{code_value}_{timestamp}.xlsx"
    print("File saved as:", filename)
    wb.save(filename)
    
    
#fills the size selection box, changes with category selected   
def gen_size_list(selected):
    item_list = []
    #selected = selected_type_category.get()
    input_box_sizes_secondary.grid_remove()2
    global has_secondary 
    has_secondary = False
    if selected == "Clothing": item_list = sizes_tops
    elif selected == "Shoes": item_list = sizes_shoes
    elif selected == "Pants": 
        item_list = sizes_pants
        has_secondary = True
        input_box_sizes_secondary.grid()
    input_box_sizes.delete(0, END)
    for item in item_list:
        input_box_sizes.insert(END, item)

#changes category between Workwear and Footwear
def change_type_category():
    cat = selected_type_category.get()
    type_label.config(text =f"Selected Category: {cat}")
    if cat == "Shoes": 
        category_label.config(text="Footwear Type") 
    else:
        category_label.config(text="Workwear Categories")
    gen_size_list(cat)

#populates colour and size lists from appropriate txt files
#this is so that users can edit the available options without requiring source code edits
def get_size_colour_lists():
    with open("./size_data/colours_listing.txt", "r", encoding="utf8") as file:
        for line in file:
            cleaned_line = line.strip()
            colours.append(cleaned_line)
    with open("./size_data/sizes_tops_listing.txt", "r", encoding="utf8") as file:
        for line in file:
            cleaned_line = line.strip()
            sizes_tops.append(cleaned_line)
    with open("./size_data/sizes_shoes_listing.txt", "r", encoding="utf8") as file:
        for line in file:
            cleaned_line = line.strip()
            sizes_shoes.append(cleaned_line)
    with open("./size_data/sizes_pants_listing.txt", "r", encoding="utf8") as file:
        for line in file:
            cleaned_line = line.strip()
            sizes_pants.append(cleaned_line)

colours = []
sizes_tops = []
sizes_shoes = []
sizes_pants = []
get_size_colour_lists()


#create the interface
root = Tk()
root.geometry("900x600")
frm = ttk.Frame(root, padding=10)
frm.grid()
root.title("Murphy Gear Webstore Item Tag Creation Tool")

#input boxes

#product ID
ttk.Label(frm, text="Product ID (No Size)").grid(column=0, row=1)
input_box_code = Entry(frm) #Text(root, width=12, height=5)
input_box_code.grid(column = 1, row = 1)

#brand
ttk.Label(frm, text="Brand").grid(column=0, row=2)
input_box_brand = Entry(frm)
input_box_brand.grid(column = 1, row = 2)

#colour selector, limits to one of predefined selection
selected_colour = StringVar(value = "Black")
ttk.Label(frm, text="Select Colour").grid(column=0, row=3)
input_box_colour = OptionMenu(frm, selected_colour, *colours)
input_box_colour.grid(column = 1, row = 3)

#type
category_label = ttk.Label(frm, text="Workwear Category")
category_label.grid(column=0, row=4)
input_box_type = Entry(frm)
input_box_type.grid(column = 1, row = 4)

selected_type_category = StringVar(value="Clothing")
Radiobutton(frm, text="Clothing", variable=selected_type_category, value="Clothing").grid(column = 2, row = 4)
Radiobutton(frm, text="Shoes", variable=selected_type_category, value="Shoes").grid(column = 3, row = 4)
Radiobutton(frm, text="Pants", variable=selected_type_category, value="Pants").grid(column = 4, row = 4)
ttk.Button(frm, text="Select Type Category", command = change_type_category).grid(column = 5, row = 4)
type_label = Label(frm, text="Selected Category: Clothing")
type_label.grid(column = 6, row = 4)

#features
features = []
def add_feature():
    feature = input_box_features.get()
    if feature not in features:
        features.append(feature)
    else:
        features.remove(feature)
    features_listing.config(text=f"Features: {features}")
    
ttk.Label(frm, text="Features").grid(column=0, row=5)
input_box_features = Entry(frm)
input_box_features.grid(column = 1, row = 5)

ttk.Button(frm, text="Add/Remove Feature", command = add_feature).grid(column = 2, row = 5)
features_listing = Label(frm, text="Features: ", justify="left", wraplength=400)
features_listing.grid(column = 3, row = 5, columnspan = 4, sticky='EW')

#sizes
ttk.Label(frm, text="Sizes").grid(column=0, row=6)
#contains currently selected sizes
selected_sizes = []
#adds a size to the current selection
#if already selected, removes instead
def add_size():
    size = input_box_sizes.get(ACTIVE)
    if has_secondary: size = size + "/" + input_box_sizes_secondary.get(ACTIVE)
    if size not in selected_sizes:
        selected_sizes.append(size)
    else:
        selected_sizes.remove(size)
    #print(selected_sizes)
    size_listing.config(text=f"Selected Sizes: {selected_sizes}")

has_secondary = False
input_box_sizes = Listbox(frm, exportselection=False)
input_box_sizes.grid(column = 1, row = 6)
input_box_sizes_secondary = Listbox(frm, exportselection=False)
input_box_sizes_secondary.grid(column = 2, row = 6)
for item in sizes_pants:
        input_box_sizes_secondary.insert(END, item)
input_box_sizes_secondary.grid_remove()
gen_size_list(selected_type_category.get())

ttk.Button(frm, text="Add/Remove Size", command = add_size).grid(column = 1, row = 7)
size_listing = Label(frm, text="Selected Sizes: ", wraplength=100, justify="left")
size_listing.grid(column = 3, row = 7)


#submit button, saves values to workbook
ttk.Button(frm, text="Submit", command=submit).grid(column=1, row=10)

#close button
ttk.Button(frm, text="Quit", command=root.destroy).grid(column=3, row=10)

root.mainloop()
