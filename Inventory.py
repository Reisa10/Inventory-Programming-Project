from openpyxl import load_workbook, Workbook

filename = "DatabaseCap.xlsx"

# Try to load the workbook, create if it doesn't exist
try:
    wb = load_workbook(filename)
    ws = wb.active
except FileNotFoundError:
    wb = Workbook()
    ws = wb.active
    ws.append(["Product Name", "Stocks"])  # headers


# Read and display stored data
def print_inv():
    print("\nCurrent Inventory:")
    for product, qty in ws.iter_rows(min_row=2, values_only=True):
        print(product, qty)


def add_new():
    fruit = input("Enter Product Name: ").title()
    qty = int(input("Enter Stocks: "))
    ws.append([fruit, qty, ])
    wb.save("DatabaseCap.xlsx")
    print(f"Added {fruit} with {qty} stocks")


def change_stock():
    product_name = input("enter product name: ").title()
    found = False

    for rows in ws.iter_rows(min_row=2):
        if rows[0].value == product_name:
            found = True
    
    if not found: 
        print("item not found")
        descicion = input("press any character to exit or A to add new product: ").upper()
        if descicion == "A":
            add_new()

    quantity = int(input("enter new quantity: "))
      
    for rows in ws.iter_rows(min_row=2):
        if rows[0].value == product_name:
            rows[1].value = quantity
            print(f"{product_name} stocks updated to {quantity}")
            wb.save("DatabaseCap.xlsx")
            found = True
            break

def remv_item():
    item = input("Enter Product to remove: ").title()
    found = False
    for row in ws.iter_rows(min_row=2):
        if row[0].value == item:
            decision = input(f"are you sure you want to remove {row[0].value}? (y/n): ").lower()
            if decision == "y":
                print(f"{row[0].value} removed")
                ws.delete_rows(row[0].row)
                found = True
                break
            else:
                print("removal cancelled")
                break
    if not found:
        d = input("item id not found, wanna try again? y/n: ").lower()
        if d == "y":
            remv_item()


while True:
    print("\nInventory Management System")
    print("1. Add New Product")
    print("2. Change Stock Quantity")
    print("3. Remove Product")
    print("4. View Inventory")
    print("5. Exit")
    choice = input("Select an option (1-5): ")

    if choice == '1':
        add_new()        
    elif choice == '2':
        change_stock()
    elif choice == '3':
        remv_item()
    elif choice == '4':
        print_inv()
    elif choice == '5':
        print("Exiting the program.")
        break
    else:
        print("Invalid choice. Please select a valid option.")

# Save the workbook
wb.save(filename)

