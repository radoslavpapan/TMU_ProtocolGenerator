import json
import os
from collections import OrderedDict

def reorder_json(data):
    # Vytvorenie nového OrderedDict s požadovaným poradím hlavných parametrov
    ordered_data = OrderedDict()
    
    # Hlavné parametre v požadovanom poradí
    main_order = [
        "UserName",
        "Start",
        "LibVersion",
        "ScrVersion",
        "AllTestsDone",
        "Passed",
        "CardTypeName",
        "SafeBytes",
        "Tests"
    ]
    
    # Usporiadanie hlavných parametrov
    for key in main_order:
        if key in data:
            ordered_data[key] = data[key]
    
    # Pridanie zostávajúcich parametrov
    for key in data:
        if key not in ordered_data:
            ordered_data[key] = data[key]
    
    # Usporiadanie parametrov v testoch
    if "Tests" in ordered_data:
        ordered_tests = []
        for test in ordered_data["Tests"]:
            ordered_test = OrderedDict()
            
            # Poradie parametrov v teste
            test_order = [
                "Code",
                "Name",
                "Passed",
                "Done",
                "Report",
                "ResultDesc"
            ]
            
            # Usporiadanie parametrov testu
            for key in test_order:
                if key in test:
                    ordered_test[key] = test[key]
            
            # Pridanie dodatočných parametrov (Min, Max, Unit)
            for key in test:
                if key not in ordered_test:
                    ordered_test[key] = test[key]
            
            ordered_tests.append(ordered_test)
        
        ordered_data["Tests"] = ordered_tests
    
    return ordered_data

def process_json_file(file_path):
    try:
        # Načítanie JSON súboru
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Pridanie verzií ak neexistujú
        data["LibVersion"] = "1.1"
        data["ScrVersion"] = "1.1"
        
        # Pridanie kódov testov
        if "Tests" in data:
            for i, test in enumerate(data["Tests"], 1):
                test["Code"] = f"T{i}"
        
        # Usporiadanie JSON
        ordered_data = reorder_json(data)
        
        # Zápis späť do súboru
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(ordered_data, f, indent=4)
            
        print(f"Súbor úspešne spracovaný: {file_path}")
        
    except json.JSONDecodeError:
        print(f"Chyba: Súbor nie je platný JSON: {file_path}")
    except Exception as e:
        print(f"Chyba pri spracovaní súboru {file_path}: {str(e)}")

def process_all_json_files():
    # Získanie zoznamu všetkých .json súborov v aktuálnom priečinku
    json_files = [f for f in os.listdir('.') if f.endswith('.json')]
    
    if not json_files:
        print("V priečinku sa nenašli žiadne JSON súbory.")
        return
    
    print(f"Nájdených {len(json_files)} JSON súborov")
    
    # Spracovanie každého súboru
    for json_file in json_files:
        process_json_file(json_file)

if __name__ == "__main__":
    print("Začínam spracovanie JSON súborov...")
    process_all_json_files()
    print("Spracovanie dokončené.")