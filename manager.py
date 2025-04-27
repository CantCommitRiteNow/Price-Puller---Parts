# price_scraper/manager.py

import os

CARS_FILE = 'cars.txt'
URLS_FILE = 'urls.txt'

def load_cars():
    if not os.path.exists(CARS_FILE):
        return []
    with open(CARS_FILE, 'r') as f:
        return [line.strip() for line in f if line.strip()]

def save_cars(cars):
    with open(CARS_FILE, 'w') as f:
        for car in cars:
            f.write(car + "\n")

def load_parts():
    if not os.path.exists(URLS_FILE):
        return []
    with open(URLS_FILE, 'r') as f:
        return [line.strip() for line in f if line.strip()]

def save_parts(parts):
    with open(URLS_FILE, 'w') as f:
        for part in parts:
            f.write(part + "\n")

def view_cars():
    cars = load_cars()
    if not cars:
        print("🚫 No cars found.")
    else:
        print("\n🚗 Cars:")
        for idx, car in enumerate(cars, 1):
            print(f"{idx}. {car}")

def add_car():
    car_name = input("🚗 Enter new car name: ").strip()
    cars = load_cars()
    if car_name in cars:
        print("⚠️ Car already exists!")
    else:
        cars.append(car_name)
        save_cars(cars)
        print(f"✅ Added car: {car_name}")

def delete_car():
    cars = load_cars()
    view_cars()
    choice = input("\nEnter car number to delete: ").strip()
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(cars):
            removed = cars.pop(idx)
            save_cars(cars)
            # Remove parts tied to that car
            parts = load_parts()
            parts = [part for part in parts if not part.startswith(removed + "|")]
            save_parts(parts)
            print(f"🗑️ Deleted car and its parts: {removed}")
        else:
            print("❌ Invalid selection.")
    except ValueError:
        print("❌ Invalid input.")

def add_part():
    cars = load_cars()
    if not cars:
        print("🚫 No cars found. Add a car first.")
        return
    view_cars()
    choice = input("\nSelect a car by number: ").strip()
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(cars):
            selected_car = cars[idx]
            url = input("🔗 Enter product URL: ").strip()
            product_name = input("🛒 Enter product name: ").strip()
            with open(URLS_FILE, 'a') as f:
                f.write(f"{selected_car}|{url}|{product_name}\n")
            print(f"✅ Added part [{product_name}] to car [{selected_car}]")
        else:
            print("❌ Invalid car selection.")
    except ValueError:
        print("❌ Invalid input.")

def view_parts():
    parts = load_parts()
    if not parts:
        print("🚫 No parts found.")
        return
    print("\n📦 Parts:")
    for idx, part in enumerate(parts, 1):
        car, url, name = part.split('|')
        print(f"{idx}. [{car}] {name} -> {url}")

def delete_part():
    parts = load_parts()
    if not parts:
        print("🚫 No parts to delete.")
        return
    view_parts()
    choice = input("\nEnter part number to delete: ").strip()
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(parts):
            removed = parts.pop(idx)
            save_parts(parts)
            print(f"🗑️ Deleted part: {removed}")
        else:
            print("❌ Invalid selection.")
    except ValueError:
        print("❌ Invalid input.")

def main_menu():
    while True:
        print("\n========== Car Parts Manager ==========")
        print("1. View Cars 🚗")
        print("2. Add Car ➕")
        print("3. Delete Car 🗑️")
        print("4. Add Part 🛒")
        print("5. View Parts 📦")
        print("6. Delete Part ❌")
        print("7. Exit 🚀")
        print("========================================")

        choice = input("Select an option (1-7): ").strip()

        if choice == '1':
            view_cars()
        elif choice == '2':
            add_car()
        elif choice == '3':
            delete_car()
        elif choice == '4':
            add_part()
        elif choice == '5':
            view_parts()
        elif choice == '6':
            delete_part()
        elif choice == '7':
            print("👋 Exiting. Bye!")
            break
        else:
            print("❌ Invalid option, try again.")

if __name__ == "__main__":
    main_menu()
