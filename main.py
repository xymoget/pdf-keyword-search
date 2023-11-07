import PyPDF2 #import packages
import openpyxl
import pandas as pd
import os

version = "1.2.0"
print(f"Developed by xymoget. Version {version}\nxymoget@gmail.com")

def extract_text(pdf_file: str):
    reader = PyPDF2.PdfReader(pdf_file) #create a reader of pdf file object
    text = ""
    for page in reader.pages: #loop over all pages of pdf
        text += page.extract_text() #extract text
    return text

def search_keywords(keyword: str, text: str):
    special_symbols = [".", "\n", ":", "-", ";", ",", "!", "?", "^"]
    for symbol in special_symbols:
        text = (" "+text).lower().replace(symbol, " ")
        keyword = keyword.lower().replace(symbol, " ")
    return text.count(f" {keyword} ") #search for keyword in the text

def load_keys(xlsx_file: str):
    wb = openpyxl.load_workbook(xlsx_file) #load excel table
    ws = wb.active
    keys = list(dict.fromkeys([str(cell.value) for cell in ws["A"]])) #get all values from A column (keywords)
    wb.close()
    return keys

def set_size(xlsx_file: str):
    wb = openpyxl.load_workbook(xlsx_file) #load excel table
    ws = wb.worksheets[0]
    for column_cells in ws.columns:
        length = max(len(as_text(cell.value)) for cell in column_cells)
        ws.column_dimensions[openpyxl.utils.get_column_letter(column_cells[0].column)].width = length
    wb.save(xlsx_file)
    wb.close()

def as_text(value):
    if value is None:
        return ""
    return str(value)
pdf_files = []
keys_path = input("Path to .xlsx file with keywords: ")
save_path = input("Path to .xlsx file to save analysis: ")
while (path:=input("Path to pdf file/directory ([-] to start analyze): ")) != "-": #inputs of data
    if os.path.isfile(path) and path.endswith(".pdf"):
        pdf_files.append(path)
    elif os.path.isdir(path):
        pdfs = [os.path.join(path, i) for i in os.listdir(path) if i.endswith(".pdf")] #loop over all files and find only pdf
        pdf_files += pdfs

pdf_files = list(dict.fromkeys(pdf_files))

keywords = load_keys(keys_path)
data = {
    "keyword": keywords
}
all_keywords = []
for file in pdf_files: #loop over all pdf files provided
    print(f"Processing file: {file}")
    short_file = file.split("\\")[-1]
    keywords_found = []
    for key in keywords: #search keywords in every file
        print(f"Searching keyword: {key}")
        text = extract_text(file)
        result = search_keywords(key, text)
        if result > 0:
            keywords_found.append(key)
        if data.get(short_file, None) != None: #save results to dictionary
            data[short_file].append(result)
        else:
            data[short_file] = [result]
    data[short_file].append("; ".join(keywords_found))
    all_keywords.extend(keywords_found)

data["keyword"].append("Keywords found")
data["keyword"].append("Total")
data["keyword"].append("Total per single keyword")
data["conclusion"] = []
print("Summing up...")
for i in range(len(data["keyword"])-3): #find files where keywords were used
    counter = 0
    files = []
    for key in data.keys():
        if key == "keyword" or key == "conclusion":
            continue
        if data[key][i] > 0:
            counter += 1
            files.append(f'"{key}"')
    conclusion = f"{str(counter)} file(s): {'; '.join(files)}"
    data["conclusion"].append(conclusion)

print(data)

for key in data.keys(): #calculate total amount of keywords used in file
    if key == "keyword" or key == "conclusion":
        continue
    data[key].append(sum(data[key][:-1]))
    data[key].append(len([i for i in data[key][:-2] if i > 0]))

unique_keywords = list(set(all_keywords))
data["conclusion"] += [f'{len(unique_keywords)} keywords found: {"; ".join(unique_keywords)}', sum([data[i][-2] for i in data if i not in ["keyword", "conclusion"]]), sum([data[i][-1] for i in data if i not in ["keyword", "conclusion"]])]

for key in data:
    temp = data[key][-3]
    data[key][-3] = data[key][-1]
    data[key][-1] = temp

df = pd.DataFrame(data) #turn dictionary into dataframe object
df.to_excel(save_path, index=False) #save to excel file
set_size(save_path)
print(f"Results are saved to {save_path}")
print(df)
print(f"Check {save_path} to see full results")
input()
