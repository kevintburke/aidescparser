from ollama import chat
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import ast
import time
import re
import datetime

class descparser:
    def __init__(self, root):
        self.root = root
        self.root.title("Item Description Parser")

        tk.Label(root, text="This program will parse item descriptions for serials and return formatted enumeration and chronology fields using a locally-hosted LLM.\nInputs must be in .csv or .xlsx and include the barcode or the MMS ID, Holdings ID, and Item PID and description for all items to create an output useable with the Update Items by Excel Cloud App").pack(pady=5)

        #Prompt user to load file
        tk.Label(root, text="Please select a file to process (in .xlsx format ONLY)").pack(pady=5)
        tk.Button(root, text="Load", command=self.getfile).pack(pady=5)

        #Ask user which ID type to use
        tk.Label(root, text="Select which form of unique item identifier to use:").pack(pady=5)
        self.usebarcode = tk.BooleanVar(value=None)
        tk.Radiobutton(text="Barcode",value=True,variable=self.usebarcode).pack(pady=5)
        tk.Radiobutton(text="MMS ID/Holdings ID, Item PID",value=False,variable=self.usebarcode).pack(pady=5)

        #Create listbox and checkbox for selecting multiple columns
        tk.Label(root, text="Select barcode column ONLY (if using barcode as ID) or MMS ID, Holdings ID, and Item PID columns (if using item ID):").pack(pady=5)
        self.list_cols = tk.Listbox(root,selectmode="multiple",exportselection=0,height=5)
        self.list_cols.pack(pady=5)

        #Create listbox for description column selecting (single selectmode)
        tk.Label(root, text="Please select a single column containing the item descriptions.").pack(pady=5)
        self.desc_col = tk.Listbox(root,exportselection=0,height=5)
        self.desc_col.pack(pady=5)

        #Run button to run full process
        tk.Label(root, text="Run process. Note: the process may take some time. See the output in the terminal for progress updates.").pack(pady=5)
        tk.Button(root, text="RUN",command=self.runparse).pack(pady=5)

        self.df = None
        self.logs = '"Call","Description Length","Duration (s)","Pass/Fail"\n'
        self.callcount = 0

    def getfile(self):
        self.list_cols.delete(0, tk.END)
        self.desc_col.delete(0, tk.END)
        file = filedialog.askopenfilename()
        #Check filetype and populate dataframe
        if file.endswith("xlsx") or file.endswith("xlsm") or file.endswith("xls"):
            try:
                self.df = pd.read_excel(file)
                print("XLSX file recognized and read")
                for column in list(self.df.columns.values):
                    self.list_cols.insert(tk.END, column)
                    self.desc_col.insert(tk.END, column)
                self.inputpresent = True
                return self.inputpresent
            except Exception as e:
                messagebox.showerror(title="Fileread error - XLSX", message=e)
                self.inputpresent = False
                return self.inputpresent
        elif file.endswith("csv"):
            try:
                self.df = pd.read_csv(file)
                print("CSV file recognized and read")
                for column in list(self.df.columns.values):
                    self.list_cols.insert(tk.END, column)
                    self.desc_col.insert(tk.END, column)
                self.inputpresent = True
                return self.inputpresent
            except Exception as e:
                messagebox.showerror(title="Fileread error - CSV", message=e)
                self.inputpresent = False
                return self.inputpresent
        else:
            messagebox.showerror(title="Filetype Error", message="Error: Please ensure file is in .xlsx or .csv format and filename includes filetype extension.")
            self.inputpresent = False
            return self.inputpresent

    def runparse(self):
        tick = time.perf_counter()
        if self.df.empty:
            messagebox.showwarning(title="Error: Missing input",detail="Please select a file to process.")
            return
        if self.usebarcode.get() == None:
            messagebox.showwarning(title="Error: Missing method",detail="Please select a method for identifying records (barcode or MMS/holdings/item ID combo).")
            return
        elif self.usebarcode.get() == True:
            method = 1
            idcols = list(self.list_cols.get(self.list_cols.curselection()))
            if len(idcols) != 1:
                messagebox.showwarning(title="Mismatch - Method and Column Count", detail="Please select exactly one column to use for identifiers if using barcodes.")
                return
        else:
            method = 0
            idcols = [self.list_cols.get(i) for i in self.list_cols.curselection()]
            #Re-order columns to ensure MMS, holdings, and item in that order
            for col in idcols:
                if str(self.df[col][0]).strip().startswith("99"):
                    mmscol = col
                elif str(self.df[col][0]).strip().startswith("22"):
                    holdingscol = col
                elif str(self.df[col][0]).strip().startswith("23"):
                    itemcol = col
                else:
                    messagebox.showwarning(title="Data Error",detail="Please ensure selected columns contain MMS IDs, Holding IDs, and Item IDs.")
                    return
            idcols = [mmscol, holdingscol, itemcol]
            if len(idcols) != 3:
                messagebox.showwarning(title="Mismatch - Method and Column Count", detail="Please select exactly three columns to use for identifiers if using barcodes.")
        desc_col = self.desc_col.get(self.desc_col.curselection())
        if desc_col == None:
            messagebox.showwarning(title="Error: Missing input",detail="Please select a column for the item descriptions to be parsed.")
            return
        dataout = []
        with open("descparser_prompt_mega.txt", "r", encoding="utf-8") as fh:
            prompt = fh.read()
        print(f'Prompt read: ',prompt)
        with open("parsecheck_prompt.txt", "r", encoding="utf-8") as fj:
            parsecheckprompt = fj.read()
        #Check that ollama is running and prompt user to activate if not, breaking function
        print("Checking ollama...")
        try:
            ticka = time.perf_counter()
            chatcheck = chat(model="qwen2.5-coder",messages=[{'role':'user','content':''}])
            print("Ollama is running:", chatcheck.message.content)
            tocka = time.perf_counter()
            self.logs += f'"{self.callcount}","-","{tocka-ticka:0.2f}","-"\n'
        except Exception as e:
            print(e)
            messagebox.showwarning(title="Ollama Error",message="Error. Please ensure ollama is running.")
            return
        for index, row in self.df.iterrows():
            self.descparse(method, desc_col, idcols, row, dataout, prompt, parsecheckprompt)
        if method == 1:
            dfout = pd.DataFrame(dataout, columns=["barcode","description","enumeration_a","enumeration_b","enumeration_c","enumeration_d","chronology_i","chronology_j","chronology_k"])
        else:
            dfout = pd.DataFrame(dataout, columns=["mms_id","holding_id","item_pid","description","enumeration_a","enumeration_b","enumeration_c","enumeration_d","chronology_i","chronology_j","chronology_k"])
        tock = time.perf_counter()
        print(f'{len(self.df)} descriptions parsed in {int(tock-tick)/60:0.2f} minutes.')
        fileout = filedialog.asksaveasfilename(defaultextension=".xlsx")
        dfout.to_excel(fileout, index=False)
        currenttime = datetime.datetime.today().strftime('%m%d%H%M')
        with open(f'{currenttime}_descparser_log.csv',"w") as fh:
            fh.write(self.logs)

    def descparse(self, method, desc_col, idcols, row, dataout, prompt, parsecheckprompt):
        #Skip empty rows
        if str(row[desc_col]) == "nan":
            print("No description found. Continuing")
            return(dataout)
        else:
            success = False
            loopcount = 1
            self.callcount+=1
            while success == False and loopcount < 4:
                tickb = time.perf_counter()
                print("Calling ollama...")
                structured_data = chat(model="qwen2.5-coder",options={"temperature":0.2, "top_p":0.9},messages=[{'role':'user','content':prompt + str(row[desc_col]),},]).message.content
                #Parse to remove ```python``` and ```json``` to reduce amount of errors and repeat calls req'd
                if structured_data.startswith("`"):
                    print("Cleaning structured data string: ", structured_data)
                    structured_data = re.sub(r'```python\n?(\{[^\}]+\})\n?```|```json\n?(\{[^\}]+\})\n?```',r'\1\2',structured_data)
                    print("Structured data cleaned to: ", structured_data)
                #Try parsing dictionary for required content
                ###REDUCE NUMBER OF FIELDS? 
                try:
                    parsed_data = ast.literal_eval(structured_data)
                    parsed_data["Enum A"]
                    parsed_data["Enum B"]
                    parsed_data["Enum C"]
                    parsed_data["Enum D"]
                    parsed_data["Chron I"]
                    parsed_data["Chron J"]
                    parsed_data["Chron K"]
                    print(f'Parse successful: {parsed_data}')
                    ###ADD ADVERSARIAL-ISH CHECK OF OUTPUT AGAINST GENERAL STRUCTURE?
                    parseconf = chat(model="qwen2.5-coder",options={"temperature":0.2, "top_p":0.9},messages=[{'role':'user','content':parsecheckprompt + str(row[desc_col]) + str(parsed_data),},]).message.content
                    if parseconf == 1:
                        success = True
                        print("Check passed!")
                        tockb = time.perf_counter()
                        self.logs += f'"{self.callcount}","{len(str(row[desc_col]))}","{tockb-tickb:0.2f}","Pass"\n'
                    else:
                        loopcount += 1
                        print("Check failed.")
                        tockb = time.perf_counter()
                        self.logs += f'"{self.callcount}","{len(str(row[desc_col]))}","{tockb-tickb:0.2f}","Check Failed"\n'
                        continue
                except:
                    print("Parse failed",structured_data)
                    tockb = time.perf_counter()
                    self.logs += f'"{self.callcount}","{len(str(row[desc_col]))}","{tockb-tickb:0.2f}","Parse Failed"\n'
                    continue
            if success == False and method == 1:
                data = {"barcode":str(row[idcols[0]]),"description":str(row[desc_col]),"enumeration_a":"FAILED","enumeration_b":"","enumeration_c":"","enumeration_d":"","chronology_i":"","chronology_j":"","chronology_k":""}
                dataout.append(data)
                return
            elif success == False and method == 0:
                data = {"mms_id":str(row[idcols[0]]),"holding_id":str(row[idcols[1]]),"item_pid":str(row[idcols[2]]),"description":str(row[desc_col]),"enumeration_a":"FAILED","enumeration_b":"","enumeration_c":"","enumeration_d":"","chronology_i":"","chronology_j":"","chronology_k":""}
                dataout.append(data)
                return
            if method == 1:
                data = {"barcode":str(row[idcols[0]]),"description":str(row[desc_col]),"enumeration_a":str(parsed_data["Enum A"]),"enumeration_b":str(parsed_data["Enum B"]),"enumeration_c":str(parsed_data["Enum C"]),"enumeration_d":str(parsed_data["Enum D"]),"chronology_i":str(parsed_data["Chron I"]),"chronology_j":str(parsed_data["Chron J"]),"chronology_k":str(parsed_data["Chron K"])}
                #rowout = pd.Series(data=data, index=["barcode","description","enumeration_a","enumeration_b","chronology_i","chronology_j"])
                dataout.append(data)
                return
            else:
                data = {"mms_id":str(row[idcols[0]]),"holding_id":str(row[idcols[1]]),"item_pid":str(row[idcols[2]]),"description":str(row[desc_col]),"enumeration_a":str(parsed_data["Enum A"]),"enumeration_b":str(parsed_data["Enum B"]),"enumeration_c":str(parsed_data["Enum C"]),"enumeration_d":str(parsed_data["Enum D"]),"chronology_i":str(parsed_data["Chron I"]),"chronology_j":str(parsed_data["Chron J"]),"chronology_k":str(parsed_data["Chron K"])}
                dataout.append(data)
                return
            
def main():
    root = tk.Tk()
    app = descparser(root)
    root.mainloop()

if __name__ == "__main__":
    main()