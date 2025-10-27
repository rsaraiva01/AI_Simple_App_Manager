import openpyxl
wb = openpyxl.Workbook()
wb = openpyxl.load_workbook("stock_fruta.xlsx")
sheetname = wb.sheetnames[0] #variável para a 1.ª folha
sheet_1 = wb[sheetname]

from fruit_synonyms import fruit_syntax, fruit_categories 

print(f"\tType 'exit' to close the chat.")
print(f"\tType 'help' for an explanation of the program.")
print(f"\n\tHello! Which fruit would you like to check in stock?")

search_fruits_stock = {}
message = ""
exit_message = ""

def search_in_message():
    message = input("You: ")
    message_prompt = message.lower().split()
    if not message:
        return
    #print(message_prompt)

    if message_prompt[0] == "exit":
        global exit_message
        exit_message = message_prompt[0]
    
    else:      
        for word in message_prompt:
            #print("Verifying if", word, "is a fruit" )
            line = 0
            max_row = sheet_1.max_row
            while line < max_row:
                search_fruits_result = sheet_1.cell(row = line + 1, column = 1).value        
                #print(search_fruits_result)
                if word == search_fruits_result:  
                    search_stock_units =  sheet_1.cell(row = line + 1, column = 2).value            
                    search_fruits_stock[search_fruits_result] = search_stock_units
                    line = max_row #Found it! Skip to search another word
                    #print("User word is a fruit:", search_fruits_stock)
                line = line + 1 
    
    show_fruits_stock()
    

    return search_fruits_stock.clear()

def show_fruits_stock():     
    if len(search_fruits_stock) > 0: 
        for name, number in search_fruits_stock.items():
            if number > 0:
                print(f"\tThe {name} has {number} unit(s) in stock.")
            else:
                print(f"\tOut of stock for {name}.")    
    else:
        print("\tHow can i help you today?")


while exit_message != "exit":
    chat = search_in_message()
    if exit_message == "exit":
        print("\tThank You!")
        break
    