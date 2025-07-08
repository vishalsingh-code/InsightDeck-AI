#!/usr/bin/env python3
"""
Interactive examples for PPT Generator
This script provides various examples and templates for different types of presentations
"""

from advanced_ppt_generator import CSVPPTGenerator

def csv_excel_data_presentation_example():
    """Generate a presentation from CSV or Excel data"""
    generator = CSVPPTGenerator()
    
    print("📊 CSV/Excel Data Presentation Generator")
    file_path = input("Enter path to your CSV or Excel file: ").strip()
    
    if not file_path:
        print("❌ No file path provided")
        return
    
    try:
        # Detect file type
        file_type = generator.detect_file_type(file_path)
        print(f"📁 Detected file type: {file_type.upper()}")
        
        # For Excel files, show sheet options
        sheet_name = None
        if file_type == 'excel':
            excel_info = generator.load_excel_info(file_path)
            print(f"\n📋 Available sheets:")
            for i, (name, info) in enumerate(excel_info['sheets'].items(), 1):
                status = "✅ Has data" if info['has_data'] else "❌ Empty"
                print(f"{i}. {name} ({info['estimated_records']} rows) - {status}")
            
            choice = input("\nSelect sheet (enter number or name, or press Enter for auto-select): ").strip()
            if choice:
                if choice.isdigit():
                    sheet_names = list(excel_info['sheets'].keys())
                    if 1 <= int(choice) <= len(sheet_names):
                        sheet_name = sheet_names[int(choice) - 1]
                else:
                    sheet_name = choice
        
        output_file = input("Enter output filename (optional): ").strip() or None
        
        print(f"\n🚀 Generating presentation from {file_type} data...")
        result = generator.create_presentation_from_csv(file_path, output_file, sheet_name)
        print(f"✅ Success! Presentation saved as: {result}")
        
    except Exception as e:
        print(f"❌ Error: {e}")

def business_presentation_example():
    """Generate a business presentation"""
    generator = CSVPPTGenerator()
    
    topics = [
        "Digital Transformation Strategy 2024",
        "Quarterly Business Review Q4 2023",
        "Market Analysis and Competitive Landscape",
        "Product Launch Strategy",
        "Team Performance and KPIs"
    ]
    
    print("🏢 Business Presentation Examples:")
    for i, topic in enumerate(topics, 1):
        print(f"{i}. {topic}")
    
    choice = input("\nSelect a topic (1-5) or enter custom topic: ").strip()
    
    if choice.isdigit() and 1 <= int(choice) <= 5:
        selected_topic = topics[int(choice) - 1]
    else:
        selected_topic = choice
    
    print(f"Generating business presentation: {selected_topic}")
    generator.generate_and_create(selected_topic, num_slides=10)

def educational_presentation_example():
    """Generate an educational presentation"""
    generator = CSVPPTGenerator()
    
    topics = [
        "Introduction to Machine Learning",
        "Climate Change and Environmental Impact",
        "World History: The Renaissance Period",
        "Mathematics: Calculus Fundamentals",
        "Biology: Cellular Structure and Function"
    ]
    
    print("🎓 Educational Presentation Examples:")
    for i, topic in enumerate(topics, 1):
        print(f"{i}. {topic}")
    
    choice = input("\nSelect a topic (1-5) or enter custom topic: ").strip()
    
    if choice.isdigit() and 1 <= int(choice) <= 5:
        selected_topic = topics[int(choice) - 1]
    else:
        selected_topic = choice
    
    print(f"Generating educational presentation: {selected_topic}")
    generator.generate_and_create(selected_topic, num_slides=12)

def technical_presentation_example():
    """Generate a technical presentation"""
    generator = CSVPPTGenerator()
    
    topics = [
        "Cloud Architecture Best Practices",
        "DevOps Implementation Strategy",
        "Cybersecurity in Modern Applications",
        "Microservices vs Monolithic Architecture",
        "API Design and RESTful Services"
    ]
    
    print("💻 Technical Presentation Examples:")
    for i, topic in enumerate(topics, 1):
        print(f"{i}. {topic}")
    
    choice = input("\nSelect a topic (1-5) or enter custom topic: ").strip()
    
    if choice.isdigit() and 1 <= int(choice) <= 5:
        selected_topic = topics[int(choice) - 1]
    else:
        selected_topic = choice
    
    print(f"Generating technical presentation: {selected_topic}")
    generator.generate_and_create(selected_topic, num_slides=15)

def custom_presentation_example():
    """Generate a custom presentation with user input"""
    generator = CSVPPTGenerator()
    
    print("📝 Custom Presentation Generator")
    topic = input("Enter your presentation topic: ").strip()
    
    while True:
        try:
            num_slides = int(input("Enter number of slides (3-20): "))
            if 3 <= num_slides <= 20:
                break
            else:
                print("Please enter a number between 3 and 20")
        except ValueError:
            print("Please enter a valid number")
    
    filename = input("Enter filename (optional, press Enter to auto-generate): ").strip()
    if not filename:
        filename = None
    
    print(f"Generating custom presentation: {topic}")
    generator.generate_and_create(topic, num_slides=num_slides, filename=filename)

def main():
    """Main interactive menu"""
    print("🎯 PPT Generator - Interactive Examples")
    print("=" * 50)
    
    while True:
        print("\nChoose presentation type:")
        print("1. 📊 CSV/Excel Data Presentation")
        print("2. 🏢 Business Presentation")
        print("3. 🎓 Educational Presentation") 
        print("4. 💻 Technical Presentation")
        print("5. 📝 Custom Presentation")
        print("6. ❌ Exit")
        
        choice = input("\nEnter your choice (1-6): ").strip()
        
        try:
            if choice == "1":
                csv_excel_data_presentation_example()
            elif choice == "2":
                business_presentation_example()
            elif choice == "3":
                educational_presentation_example()
            elif choice == "4":
                technical_presentation_example()
            elif choice == "5":
                custom_presentation_example()
            elif choice == "6":
                print("👋 Goodbye!")
                break
            else:
                print("❌ Invalid choice. Please enter 1-6.")
                continue
                
            print("\n" + "="*50)
            
        except KeyboardInterrupt:
            print("\n👋 Goodbye!")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            continue

if __name__ == "__main__":
    main()
