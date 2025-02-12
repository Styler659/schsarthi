import pandas as pd
import os

class DatabaseManager:
    def __init__(self, filename="database.xlsm"):
        base_dir = os.path.abspath(os.path.dirname(__file__))
        self.filepath = os.path.join(base_dir, filename)
        self.load_database()

    def load_database(self):
        """Loads the database from the Excel file."""
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"Database file '{self.filepath}' not found.")
        
        try:
            self.df = pd.read_excel(self.filepath, engine="openpyxl")
        except Exception as e:
            raise Exception(f"Error loading database: {e}")

    def save_database(self):
        """Saves the current DataFrame back to the Excel file."""
        try:
            self.df.to_excel(self.filepath, index=False, engine="openpyxl")
        except Exception as e:
            raise Exception(f"Error saving database: {e}")

    def add_scheme(self, scheme_data):
        """Adds a new scheme to the database, prompting for age limits if necessary."""
        if scheme_data.get("age group") == "S":
            scheme_data["ageL"] = input("Enter lower age limit: ")
            scheme_data["ageH"] = input("Enter upper age limit: ")
        
        new_entry = pd.DataFrame([scheme_data])
        self.df = pd.concat([self.df, new_entry], ignore_index=True)
        self.save_database()

    def edit_scheme(self, scheme_name, updated_data):
        """Edits an existing scheme based on its name, updating all fields."""
        index = self.df[self.df["Scheme name"] == scheme_name].index
        if index.empty:
            raise ValueError("Scheme not found.")
        
        for col in self.df.columns:
            if col in updated_data:
                self.df.at[index[0], col] = updated_data[col]
        
        self.save_database()

    def delete_scheme(self, scheme_name):
        """Deletes a scheme from the database based on its name."""
        index = self.df[self.df["Scheme name"] == scheme_name].index
        if index.empty:
            raise ValueError("Scheme not found.")
        
        self.df.drop(index, inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.save_database()

if __name__ == "__main__":
    db_manager = DatabaseManager()
    
    while True:
        print("\nDatabase Manager")
        print("1. Add Scheme")
        print("2. Edit Scheme")
        print("3. Delete Scheme")
        print("4. Exit")
        
        choice = input("Enter your choice: ").strip()
        
        if choice == "1":
            scheme_data = {
                "Scheme name": input("Enter scheme name: "),
                "Info": input("Enter scheme info: "),
                "link": input("Enter scheme link: "),
                "farmer": input("Is it for farmers? (Y/N): ").strip().upper(),
                "rural": input("Is it for rural areas? (Y/N): ").strip().upper(),
                "pregnant": input("Is it for pregnant women? (Y/N): ").strip().upper(),
                "income": input("Income category (ALL/BPL/LIG/EWS/ABOVE): ").strip().upper(),
                "gender": input("Gender restriction (M/F/B): ").strip().upper(),
                "age group": input("Age group (A/STUDENT/WA/SENIOR/S): ").strip().upper(),
                "pwd": input("Is it for disabled people? (Y/N/B): ").strip().upper(),
            }
            
            if scheme_data["age group"] == "S":
                scheme_data["ageL"] = input("Enter lower age limit: ")
                scheme_data["ageH"] = input("Enter upper age limit: ")
            
            db_manager.add_scheme(scheme_data)
            print("Scheme added successfully!")

        elif choice == "2":
            scheme_name = input("Enter scheme name to edit: ")
            updated_data = {
                "Scheme name": scheme_name,
                "Info": input("Enter updated info: "),
                "link": input("Enter updated link: "),
                "farmer": input("Is it for farmers? (Y/N): ").strip().upper(),
                "rural": input("Is it for rural areas? (Y/N): ").strip().upper(),
                "pregnant": input("Is it for pregnant women? (Y/N): ").strip().upper(),
                "income": input("Income category (ALL/BPL/LIG/EWS/ABOVE): ").strip().upper(),
                "gender": input("Gender restriction (M/F/B): ").strip().upper(),
                "age group": input("Age group (A/STUDENT/WA/SENIOR/S): ").strip().upper(),
                "pwd": input("Is it for disabled people? (Y/N/B): ").strip().upper(),
            }
            
            if updated_data["age group"] == "S":
                updated_data["ageL"] = input("Enter new lower age limit: ")
                updated_data["ageH"] = input("Enter new upper age limit: ")
            
            db_manager.edit_scheme(scheme_name, updated_data)
            print("Scheme updated successfully!")

        elif choice == "3":
            scheme_name = input("Enter scheme name to delete: ")
            db_manager.delete_scheme(scheme_name)
            print("Scheme deleted successfully!")

        elif choice == "4":
            print("Exiting program.")
            break

        else:
            print("Invalid choice. Please try again.")
