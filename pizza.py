# Isaac Kremer
# CS 115
# 11/5/2024
# Pizza Shop: Tracks the pizza shops orders, inventory and financials
# while allowing users to restock the inventory, order pizzas, and check
# on finances. Full stock is 10 in inventory and Account Balance begins at $1,000.00.

from Library.functions import order
from Library.functions import subtract_inventory
from Library.functions import sales
from Library.functions import view_inventory
from Library.functions import restock
from Library.functions import view_pl
from Library.functions import balance
from Library.functions import writetoerrorlog

def main():
    while True:
        try:
            print()
            print("Please select a menu item.")
            print("1. Order a Pizza")
            print("2. View inventory levels")
            print("3. Restock inventory (for shop managers)")
            print("4. View the Profit and Loss statement")
            print("5. View the current bank balance.")
            print("6. Exit the program")

            choice = input("=>")

            if choice == "1":
                customer, total, top_lst, size = order()
                subtract_inventory(top_lst, size)
                sales(customer, total)
            elif choice == "2":
                view_inventory()
            elif choice == "3":
                restock()
            elif choice == "4":
                view_pl()
            elif choice == "5":
                balance()
            elif choice == "6":
                break
            else:
                print("Please select a valid menu item.")
        except Exception as e:
            print("Error")
            writetoerrorlog("pizza", e)

if __name__ == "__main__":
    main()

    