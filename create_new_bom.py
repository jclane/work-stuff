import csv
import tkinter as tk  
    
    
PART_CLASSES = {
                "Desktop":
                    {
                    "BRD":{"MB":"Motherboard"}, 
                    "COS":{"CHASSIS":"Chassis"}, 
                    "FAN":{"FAN":"Fan"},
                    "HEAT SYNC":{"HS":"Heat sink"},
                    "PWR":{"PWR SUPPLY":"Power supply"}, 
                    "VBRD":{"GPX":"Graphics card"},
                    },
                "Laptop":
                    {
                    "AC":{"AC":"AC adaptor"}, 
                    "ANT":{"ANT":"WIFI antenna"}, 
                    "AUDIO":{"SPK":"Speakers"}, 
                    "BAT":{"BAT":"Battery"}, 
                    "BRA":{"BRA":"Hinge brackets", "HG":"Hinges"}, 
                    "BRD":{"MB":"Motherboard"},
                    "CAB":{"LCDH":"LCD cable harness"}, 
                    "COS":{"BC":"Bottom cover", "TC":"Top cover", "RC":"LCD rear cover"}, 
                    "CAM":{"CAM":"Webcam"}, 
                    "KB":{"KB":"Keyboard"}, 
                    "LCD":{"LCD":"LCD display"}, 
                    "WIR":{"WIR":"WIFI card"}
                    }
                }
                
BRAND_SUFFIX = {"brand 1":"IBY", "brand 2":"MSS", "brand 3":"RZR", "brand 4":"CBP"}

class Main(tk.Tk):
    
    def __init__(self, *args, **kwargs):
        
        def create_bom():
            model = model_entry.get().replace("-"," ").replace("\n", "").strip()
            brand = brandvar.get().lower()
            lapordesk = typevar.get()
            rows = []
            
            def csv_writer(file, rows):
                
                with open(file, 'w', newline='', encoding='latin1') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(rows)

            for part_class in PART_CLASSES[lapordesk]:

                classes = PART_CLASSES[lapordesk]
                class_actual = part_class
                classes_generic = classes[class_actual]
                for genericclass in classes_generic:
                    class_description = classes_generic[genericclass]
                    rows.append([class_actual, model.replace(" ", "") + " " + genericclass, class_description + " for a " + model_entry.get().strip()])
             
            csv_writer(model + "-" + BRAND_SUFFIX[brand].upper() + ".csv", rows)
                           
        tk.Tk.__init__(self, *args, **kwargs)
        
        window = tk.Frame(self)
        window.grid(row=0, column=0, padx=10, pady=10)
        
        typevar = tk.StringVar()
        
        type_label = tk.Label(window, text="Select Type:")
        type_drop = tk.OptionMenu(window, typevar, "Desktop", "Laptop")
        typevar.set("Desktop")
        
        model_label = tk.Label(window, text="Enter Model Number:")
        model_entry = tk.Entry(window)
        
        brandvar = tk.StringVar()
        brands = ("brand 1", "brand 2", "brand 3", "brand 4")
        
        brand_label = tk.Label(window, text="Select Brand:")
        brand_drop = tk.OptionMenu(window, brandvar, *brands)
        brandvar.set("brand 1")
            
        button = tk.Button(window, text="Create BOM", command=create_bom)
                
        type_label.grid(row=0, column=0, pady=5, sticky='e')
        type_drop.grid(row=0, column=1, pady=5, sticky='ew')
        model_label.grid(row=2, column=0, pady=5, sticky='e')
        model_entry.grid(row=2, column=1, pady=5, sticky='e')
        brand_label.grid(row=3, column=0, pady=5, sticky='e')
        brand_drop.grid(row=3, column=1, pady=5, sticky='ew')
        button.grid(row=4, column=1, padx=10, pady=5, sticky='ew')

        
if __name__ == "__main__":
    app = Main()
    app.title("Generate Secondary Brand BOM")
    app.mainloop()
