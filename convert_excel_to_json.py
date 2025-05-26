import pandas as pd
import json
import re
import os
from datetime import datetime

def assign_activity_type(activity, notes):
    """Assign activity type(s) based on activity and notes."""
    activity_lower = activity.lower()
    notes_lower = notes.lower()
    types = []

    # Travel: flights, taxis, trains, water buses, walks to transport, check-in/out, luggage
    if any(keyword in activity_lower for keyword in [
        'arrive', 'flight', 'taxi', 'walk to', 'vaporetto', 'train', 'water bus',
        'check-in', 'check out', 'luggage', 'customs', 'immigration'
    ]) or any(keyword in notes_lower for keyword in ['flight', 'taxi', 'vaporetto', 'train']):
        types.append('travel')

    # Dining: meals (breakfast, lunch, dinner, gelato, snacks)
    if any(keyword in activity_lower for keyword in [
        'breakfast', 'lunch', 'dinner', 'gelato', 'snack'
    ]):
        types.append('dining')

    # Activity: sightseeing, free time, gondola rides, markets
    if any(keyword in activity_lower for keyword in [
        'view', 'relax', 'explore', 'free time', 'gondola', 'market'
    ]) or any(keyword in notes_lower for keyword in ['photo', 'sightseeing', 'market']):
        types.append('activity')

    return types if types else ['activity']

def process_notes(notes):
    """Split notes into costs_and_notes and tickets_to_buy, remove 'Hold kids’ hands'."""
    if pd.isna(notes):
        return [], ["None."]
    
    # Remove 'Hold kids’ hands', 'watch kids', and variants
    notes = re.sub(r'(?:[Hh]old\s+(?:kids’?|children’s?)\s+hands?|watch\s+(?:kids|children))', '', notes, flags=re.IGNORECASE).strip()
    if not notes:
        return [], ["None."]
    
    # Split notes into sentences, handling multiple separators
    sentences = [s.strip() for s in re.split(r'[.\n;]', notes) if s.strip()]
    
    # Separate costs and tickets
    costs_and_notes = []
    tickets_to_buy = []
    
    for sentence in sentences:
        # Identify cost-related notes
        if 'cost' in sentence.lower() or '€' in sentence or any(keyword in sentence.lower() for keyword in [
            'apple pay', 'cash', 'ticket', 'fare', 'price', 'booking', 'reservation'
        ]):
            # If it mentions a specific cost, add to tickets_to_buy
            if '€' in sentence or any(keyword in sentence.lower() for keyword in ['fare', 'ticket']):
                tickets_to_buy.append(sentence)
            costs_and_notes.append(sentence)
        else:
            costs_and_notes.append(sentence)
    
    return costs_and_notes if costs_and_notes else [], tickets_to_buy if tickets_to_buy else ["None."]

def generate_map_link(location):
    """Generate placeholder Google Maps link(s) for the location."""
    locations = [loc.strip() for loc in re.split(r'\s*(?:to|→)\s*', location)]
    if len(locations) > 1:
        return [f"https://maps.app.goo.gl/{loc.replace(' ', '').replace('(', '').replace(')', '')}" for loc in locations]
    return f"https://maps.app.goo.gl/{location.replace(' ', '').replace('(', '').replace(')', '')}"

def convert_date(date_str):
    """Convert date from '3-Jul-25' or 'YYYY-MM-DD HH:MM:SS' to 'July 3, 2025'."""
    try:
        # Handle '3-Jul-25' format
        date_obj = datetime.strptime(date_str, '%d-%b-%y')
        return date_obj.strftime('%B %d, %Y').replace(' 0', ' ')
    except ValueError:
        try:
            # Handle '2025-07-03 00:00:00' format
            date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
            return date_obj.strftime('%B %d, %Y').replace(' 0', ' ')
        except ValueError as e:
            print(f"Error parsing date '{date_str}': {e}")
            return date_str

def normalize_columns(columns):
    """Normalize column names by stripping whitespace and standardizing case."""
    return [col.strip().lower() if isinstance(col, str) else str(col).lower() for col in columns]

def table_to_json(excel_file, header_row=0, skip_rows=0):
    """Convert Excel tabs to JSON, grouping activities by date across cities."""
    xls = pd.ExcelFile(excel_file)
    date_to_activities = {}  # Map date to list of activities
    expected_columns = ['date', 'time', 'activity', 'location', 'directions', 'transportation details', 'notes']

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=header_row, skiprows=skip_rows)
        
        # Debug: Print first 5 rows and column names
        print(f"\nSheet '{sheet_name}' first 5 rows:")
        print(df.head().to_string())
        actual_columns = normalize_columns(df.columns)
        print(f"Sheet '{sheet_name}' columns (normalized): {actual_columns}")

        # Validate columns
        if not all(col in actual_columns for col in expected_columns):
            missing = [col for col in expected_columns if col not in actual_columns]
            raise ValueError(f"Sheet {sheet_name} missing required columns: {missing}. Found: {df.columns.tolist()}")

        # Map normalized columns
        column_map = {col.strip().lower(): col for col in df.columns}
        df.columns = [col.strip().lower() for col in df.columns]

        for _, row in df.iterrows():
            date = str(row['date'])
            formatted_date = convert_date(date)
            activity = str(row['activity']) if not pd.isna(row['activity']) else ''
            location = str(row['location']) if not pd.isna(row['location']) else ''
            directions = str(row['directions']) if not pd.isna(row['directions']) else 'N/A'
            transportation = str(row['transportation details']) if not pd.isna(row['transportation details']) else 'N/A'
            notes = str(row['notes']) if not pd.isna(row['notes']) else ''

            directions_list = [d.strip() for d in directions.split('.') if d.strip()] if directions != 'N/A' else ['N/A']
            transportation_list = [t.strip() for t in transportation.split('.') if t.strip()] if transportation != 'N/A' else ['N/A']
            costs_and_notes, tickets_to_buy = process_notes(notes)

            things_to_do = []
            activity_lower = activity.lower()
            if 'arrive' in activity_lower:
                things_to_do.append(f"Arrive at {location}.")
            if 'walk' in activity_lower:
                things_to_do.append(f"Walk to {location}.")
            if 'taxi' in activity_lower or 'flight' in activity_lower or 'train' in activity_lower or 'vaporetto' in activity_lower:
                things_to_do.append(f"Travel to {location}.")
            if 'breakfast' in activity_lower or 'lunch' in activity_lower or 'dinner' in activity_lower or 'gelato' in activity_lower:
                things_to_do.append(f"Enjoy {activity_lower} at {location}.")
            if 'view' in activity_lower or 'explore' in activity_lower or 'relax' in activity_lower or 'free time' in activity_lower:
                things_to_do.append(f"Explore or relax at {location}.")
            if 'check-in' in activity_lower or 'check out' in activity_lower:
                things_to_do.append(f"Complete check-in/out at {location}.")
            if 'luggage' in activity_lower:
                things_to_do.append(f"Handle luggage at {location}.")
            if not things_to_do:
                things_to_do.append(f"Perform activity: {activity}.")

            activity_data = {
                "time": str(row['time']),
                "activity": activity,
                "location": location,
                "mapLink": generate_map_link(location),
                "directions": directions_list,
                "transportation": transportation_list,
                "thingsToDo": things_to_do,
                "ticketsToBuy": tickets_to_buy,
                "costsAndNotes": costs_and_notes,
                "type": assign_activity_type(activity, notes),
                "city": sheet_name  # Add city to activity for context
            }

            if formatted_date not in date_to_activities:
                date_to_activities[formatted_date] = []
            date_to_activities[formatted_date].append(activity_data)

    # Create itinerary by date
    itinerary = []
    for date, activities in sorted(date_to_activities.items()):
        print(f"Date '{date}' has {len(activities)} activities")
        itinerary.append({
            "city": ", ".join(sorted(set(a["city"] for a in activities))),  # List all cities for the date
            "date": date,
            "activities": activities
        })

    return {"itinerary": itinerary}

def update_html(html_file, json_data):
    """Update the itineraryData section in the HTML file."""
    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Debug: Print JSON to verify
        json_str = json.dumps(json_data, indent=2, ensure_ascii=False)
        print("\nGenerated JSON (first 500 characters):")
        print(json_str[:500] + "..." if len(json_str) > 500 else json_str)

        # Find and replace the itineraryData
        pattern = r'const itineraryData = \{[\s\S]*?\};'
        new_json_str = f"const itineraryData = {json_str};"
        updated_html = re.sub(pattern, new_json_str, html_content)

        # Write back to the file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(updated_html)
        print(f"Successfully updated {html_file}")
    except Exception as e:
        print(f"Error updating HTML: {e}")
        # Save JSON to file for inspection
        with open('debug_output.json', 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        print("Saved JSON to 'debug_output.json' for inspection")

def main():
    # Specify paths
    excel_file = "/Users/family/Dropbox/Europe/Table.xlsx"
    html_file = "/Users/family/Dropbox/Europe/itinerary.html"

    # Validate file existence
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Excel file not found: {excel_file}")
    if not os.path.exists(html_file):
        raise FileNotFoundError(f"HTML file not found: {html_file}")

    # Convert table to JSON
    try:
        json_data = table_to_json(excel_file, header_row=0, skip_rows=0)
    except ValueError as e:
        print(f"Initial attempt failed: {e}")
        print("Trying with header_row=1, skip_rows=0...")
        try:
            json_data = table_to_json(excel_file, header_row=1, skip_rows=0)
        except ValueError as e:
            print(f"Second attempt failed: {e}")
            raise

    # Update HTML
    update_html(html_file, json_data)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")