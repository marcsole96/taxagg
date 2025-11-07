import pandas as pd
import os
import re
import xlsxwriter

# Use current directory where the script is located
input_dir = os.path.dirname(os.path.abspath(__file__)) or '.'
output_file = os.path.join(input_dir, 'merged_data_with_charts.xlsx')

# Get all Excel files in the directory, excluding any output files
excel_files = sorted([f for f in os.listdir(input_dir) 
                      if f.endswith('.xlsx') and not f.startswith('merged_data')])

print(f"Found {len(excel_files)} Excel files to merge")
print(f"Working directory: {input_dir}\n")

# Initialize lists to store processed data
demographics_data = []
usability_data = []

# Response mapping for Usability questions
response_mapping = {
    'Strongly Agree (5)': 5,
    'Agree (4)': 4,
    'Neutral (3)': 3,
    'Disagree (2)': 2,
    'Strongly Disagree (1)': 1,
    'Not applicable': 0,
    'Not applicable ': 0
}

# Dictionary to store question texts (extract from first file)
question_texts = {}

# Process each Excel file
for file_idx, file_name in enumerate(excel_files):
    file_path = os.path.join(input_dir, file_name)
    participant_name = file_name.replace('.xlsx', '')
    
    print(f"Processing: {file_name}")
    
    try:
        # ===== Process Demographics Sheet =====
        df_demo = pd.read_excel(file_path, sheet_name='Demographics')
        
        demo_dict = {'Participant': participant_name}
        
        # Handle different column name formats
        question_col = 'Question' if 'Question' in df_demo.columns else df_demo.columns[0]
        
        for _, row in df_demo.iterrows():
            question = str(row[question_col]).strip()
            answer = row['Answer']
            demo_dict[question] = answer
        
        demographics_data.append(demo_dict)
        
        # ===== Process Usability Sheet =====
        df_usability = pd.read_excel(file_path, sheet_name='Usability')
        usability_dict = {'Participant': participant_name}
        headers = df_usability.iloc[2].tolist()
        
        # Extract question texts from first file only
        if file_idx == 0:
            for idx in range(3, min(21, len(df_usability))):
                row = df_usability.iloc[idx]
                question_text = str(row.iloc[0]).strip()
                
                q_match = re.match(r'(Q\d+)\)', question_text)
                if q_match:
                    question_num = q_match.group(1)
                    question_texts[question_num] = question_text
        
        for idx in range(3, min(21, len(df_usability))):
            row = df_usability.iloc[idx]
            question_text = str(row.iloc[0]).strip()
            
            q_match = re.match(r'(Q\d+)\)', question_text)
            if not q_match:
                continue
            
            question_num = q_match.group(1)
            response_value = None
            response_text = None
            
            # Priority: look for x marks first
            for col_idx in range(1, min(7, len(row))):
                cell_value = str(row.iloc[col_idx]).strip().lower()
                if 'x' in cell_value:
                    header = str(headers[col_idx]).strip()
                    response_text = header
                    response_value = response_mapping.get(header, None)
                    break
            
            # If no x, look for ( )
            if response_value is None:
                for col_idx in range(1, min(7, len(row))):
                    cell_value = str(row.iloc[col_idx]).strip()
                    if cell_value in ['( )', '()']:
                        header = str(headers[col_idx]).strip()
                        response_text = header
                        response_value = response_mapping.get(header, None)
                        break
            
            usability_dict[f'{question_num}_Score'] = response_value
            usability_dict[f'{question_num}_Response'] = response_text
        
        usability_data.append(usability_dict)
        
    except Exception as e:
        print(f"  Error: {e}")

demographics_wide = pd.DataFrame(demographics_data)
usability_wide = pd.DataFrame(usability_data)

print(f"\nExtracted {len(question_texts)} question texts")

# ===== CREATE SUMMARY SHEETS WITH FULL QUESTION TEXT =====
print("\nCreating summary sheets...")

# Demographics Summary
demo_summary_data = []

if 'Q2) What is your gender?' in demographics_wide.columns:
    gender_counts = demographics_wide['Q2) What is your gender?'].value_counts()
    for gender, count in gender_counts.items():
        demo_summary_data.append({
            'Question': 'Q2) What is your gender?',
            'Short_Name': 'Q2) Gender',
            'Response': gender,
            'Count': count,
            'Percentage': round(count/len(demographics_wide)*100, 1)
        })

if 'Q7) Which country you feel most connected to? This may not be the country where you were born' in demographics_wide.columns:
    country_col = 'Q7) Which country you feel most connected to? This may not be the country where you were born'
    country_counts = demographics_wide[country_col].value_counts()
    for country, count in country_counts.items():
        demo_summary_data.append({
            'Question': country_col,
            'Short_Name': 'Q7) Country',
            'Response': country,
            'Count': count,
            'Percentage': round(count/len(demographics_wide)*100, 1)
        })

if 'Q3) What is your most recent degree?  (e.g. BSc in Electrical Engineering)' in demographics_wide.columns:
    degree_col = 'Q3) What is your most recent degree?  (e.g. BSc in Electrical Engineering)'
    degree_counts = demographics_wide[degree_col].value_counts()
    for degree, count in degree_counts.items():
        demo_summary_data.append({
            'Question': degree_col,
            'Short_Name': 'Q3) Degree',
            'Response': degree,
            'Count': count,
            'Percentage': round(count/len(demographics_wide)*100, 1)
        })

if 'Q4) Have you ever used GenerativeAI (GenAI)? (Yes/No)?' in demographics_wide.columns:
    genai_counts = demographics_wide['Q4) Have you ever used GenerativeAI (GenAI)? (Yes/No)?'].value_counts()
    for response, count in genai_counts.items():
        demo_summary_data.append({
            'Question': 'Q4) Have you ever used GenerativeAI (GenAI)? (Yes/No)?',
            'Short_Name': 'Q4) Used GenAI',
            'Response': response,
            'Count': count,
            'Percentage': round(count/len(demographics_wide)*100, 1)
        })

if 'Q5) If you answered \'Yes\' to Q4, how often do you use GenAI? (e.g. once a week)' in demographics_wide.columns:
    freq_col = 'Q5) If you answered \'Yes\' to Q4, how often do you use GenAI? (e.g. once a week)'
    freq_counts = demographics_wide[freq_col].value_counts()
    for freq, count in freq_counts.items():
        demo_summary_data.append({
            'Question': freq_col,
            'Short_Name': 'Q5) GenAI Frequency',
            'Response': freq,
            'Count': count,
            'Percentage': round(count/len(demographics_wide)*100, 1)
        })

demographics_summary = pd.DataFrame(demo_summary_data)

# Usability Summary - WITH FULL QUESTION TEXT
usability_summary_data = []
for q_num in range(1, 19):
    score_col = f'Q{q_num}_Score'
    response_col = f'Q{q_num}_Response'
    
    if score_col in usability_wide.columns:
        response_counts = usability_wide[response_col].value_counts()
        
        for response, count in response_counts.items():
            if pd.notna(response):
                usability_summary_data.append({
                    'Question_Number': f'Q{q_num}',
                    'Question_Text': question_texts.get(f'Q{q_num}', f'Q{q_num}'),
                    'Response': response,
                    'Count': count,
                    'Percentage': round(count/len(usability_wide)*100, 1)
                })

usability_summary = pd.DataFrame(usability_summary_data)

# Usability Average Scores - WITH FULL QUESTION TEXT
usability_avg_data = []
for q_num in range(1, 19):
    score_col = f'Q{q_num}_Score'
    if score_col in usability_wide.columns:
        avg_score = usability_wide[score_col].mean()
        usability_avg_data.append({
            'Question_Number': f'Q{q_num}',
            'Question_Text': question_texts.get(f'Q{q_num}', f'Q{q_num}'),
            'Average_Score': round(avg_score, 2),
            'Responses': usability_wide[score_col].notna().sum()
        })

usability_averages = pd.DataFrame(usability_avg_data)

# ===== WRITE TO EXCEL WITH CHARTS =====
print(f"\nWriting to Excel with embedded charts: {output_file}")

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# Write data to sheets
demographics_wide.to_excel(writer, sheet_name='Demographics', index=False)
usability_wide.to_excel(writer, sheet_name='Usability', index=False)
demographics_summary.to_excel(writer, sheet_name='Demo_Summary', index=False)
usability_summary.to_excel(writer, sheet_name='Usability_Summary', index=False)
usability_averages.to_excel(writer, sheet_name='Usability_Averages', index=False)

print("Creating charts...")

# ===== CHART 1: COUNTRY DISTRIBUTION =====
chart_sheet = workbook.add_worksheet('Charts_Demographics')
chart_sheet.set_column('A:A', 2)

country_data = demographics_summary[demographics_summary['Short_Name'] == 'Q7) Country']
if not country_data.empty:
    # Write data for country chart
    chart_sheet.write_row('B2', ['Country', 'Count', 'Percentage'])
    for i, row in enumerate(country_data.itertuples(), start=3):
        chart_sheet.write_row(f'B{i}', [row.Response, row.Count, row.Percentage])
    
    # Create bar chart
    chart1 = workbook.add_chart({'type': 'bar'})
    chart1.add_series({
        'name': 'Participant Count',
        'categories': f'=Charts_Demographics!$B$3:$B${3+len(country_data)-1}',
        'values': f'=Charts_Demographics!$C$3:$C${3+len(country_data)-1}',
        'data_labels': {'value': True},
    })
    chart1.set_title({'name': 'Country Distribution'})
    chart1.set_x_axis({'name': 'Number of Participants'})
    chart1.set_y_axis({'name': 'Country'})
    chart1.set_size({'width': 720, 'height': 480})
    chart_sheet.insert_chart('B10', chart1)
    print("  ✓ Country distribution chart")

# ===== CHART 2: GENDER DISTRIBUTION =====
gender_data = demographics_summary[demographics_summary['Short_Name'] == 'Q2) Gender']
if not gender_data.empty:
    # Pie chart
    chart2 = workbook.add_chart({'type': 'pie'})
    
    # Write data
    start_row = 3 + len(country_data) + 5
    chart_sheet.write_row(f'B{start_row}', ['Gender', 'Count'])
    for i, row in enumerate(gender_data.itertuples(), start=start_row+1):
        chart_sheet.write_row(f'B{i}', [row.Response, row.Count])
    
    chart2.add_series({
        'name': 'Gender Distribution',
        'categories': f'=Charts_Demographics!$B${start_row+1}:$B${start_row+len(gender_data)}',
        'values': f'=Charts_Demographics!$C${start_row+1}:$C${start_row+len(gender_data)}',
        'data_labels': {'percentage': True, 'category': True},
    })
    chart2.set_title({'name': 'Gender Distribution'})
    chart2.set_size({'width': 480, 'height': 400})
    chart_sheet.insert_chart('J10', chart2)
    print("  ✓ Gender distribution chart")

# ===== CHART 3: GENAI FREQUENCY =====
freq_data = demographics_summary[demographics_summary['Short_Name'] == 'Q5) GenAI Frequency']
if not freq_data.empty:
    start_row = start_row + len(gender_data) + 5
    chart_sheet.write_row(f'B{start_row}', ['Frequency', 'Count'])
    for i, row in enumerate(freq_data.itertuples(), start=start_row+1):
        chart_sheet.write_row(f'B{i}', [row.Response, row.Count])
    
    chart3 = workbook.add_chart({'type': 'column'})
    chart3.add_series({
        'name': 'Frequency Count',
        'categories': f'=Charts_Demographics!$B${start_row+1}:$B${start_row+len(freq_data)}',
        'values': f'=Charts_Demographics!$C${start_row+1}:$C${start_row+len(freq_data)}',
        'data_labels': {'value': True},
    })
    chart3.set_title({'name': 'GenAI Usage Frequency'})
    chart3.set_x_axis({'name': 'Frequency'})
    chart3.set_y_axis({'name': 'Number of Participants'})
    chart3.set_size({'width': 640, 'height': 400})
    chart_sheet.insert_chart('B35', chart3)
    print("  ✓ GenAI frequency chart")

# ===== CHART 4: USABILITY AVERAGE SCORES =====
chart_sheet2 = workbook.add_worksheet('Charts_Usability')
chart_sheet2.set_column('A:A', 2)

# Write data
chart_sheet2.write_row('B2', ['Question', 'Average Score'])
for i, row in enumerate(usability_averages.itertuples(), start=3):
    chart_sheet2.write_row(f'B{i}', [row.Question_Number, row.Average_Score])

chart4 = workbook.add_chart({'type': 'bar'})
chart4.add_series({
    'name': 'Average Score',
    'categories': f'=Charts_Usability!$B$3:$B${3+len(usability_averages)-1}',
    'values': f'=Charts_Usability!$C$3:$C${3+len(usability_averages)-1}',
    'data_labels': {'value': True, 'num_format': '0.00'},
})
chart4.set_title({'name': 'Average Usability Scores (Q1-Q18)'})
chart4.set_x_axis({'name': 'Average Score (1-5 scale)', 'min': 0, 'max': 5})
chart4.set_y_axis({'name': 'Question'})
chart4.set_size({'width': 720, 'height': 600})
chart_sheet2.insert_chart('B25', chart4)
print("  ✓ Usability average scores chart")

# ===== CHART 5: TOP 5 & BOTTOM 5 =====
top5 = usability_averages.nlargest(5, 'Average_Score')
bottom5 = usability_averages.nsmallest(5, 'Average_Score')

# Write top 5
start_row = 3 + len(usability_averages) + 5
chart_sheet2.write_row(f'B{start_row}', ['Top 5 Questions', 'Score'])
for i, row in enumerate(top5.itertuples(), start=start_row+1):
    chart_sheet2.write_row(f'B{i}', [row.Question_Number, row.Average_Score])

chart5 = workbook.add_chart({'type': 'bar'})
chart5.add_series({
    'name': 'Top 5 Highest Scores',
    'categories': f'=Charts_Usability!$B${start_row+1}:$B${start_row+5}',
    'values': f'=Charts_Usability!$C${start_row+1}:$C${start_row+5}',
    'data_labels': {'value': True, 'num_format': '0.00'},
    'fill': {'color': '#2ecc71'},
})
chart5.set_title({'name': 'Top 5 Highest Rated Questions'})
chart5.set_x_axis({'name': 'Average Score', 'min': 0, 'max': 5})
chart5.set_size({'width': 600, 'height': 400})
chart_sheet2.insert_chart('J2', chart5)
print("  ✓ Top 5 questions chart")

# Write bottom 5
start_row = start_row + 10
chart_sheet2.write_row(f'B{start_row}', ['Bottom 5 Questions', 'Score'])
for i, row in enumerate(bottom5.itertuples(), start=start_row+1):
    chart_sheet2.write_row(f'B{i}', [row.Question_Number, row.Average_Score])

chart6 = workbook.add_chart({'type': 'bar'})
chart6.add_series({
    'name': 'Bottom 5 Lowest Scores',
    'categories': f'=Charts_Usability!$B${start_row+1}:$B${start_row+5}',
    'values': f'=Charts_Usability!$C${start_row+1}:$C${start_row+5}',
    'data_labels': {'value': True, 'num_format': '0.00'},
    'fill': {'color': '#e74c3c'},
})
chart6.set_title({'name': 'Bottom 5 Lowest Rated Questions'})
chart6.set_x_axis({'name': 'Average Score', 'min': 0, 'max': 5})
chart6.set_size({'width': 600, 'height': 400})
chart_sheet2.insert_chart('J25', chart6)
print("  ✓ Bottom 5 questions chart")

# ===== CHARTS: INDIVIDUAL QUESTIONS Q1-Q18 WITH IN-CHART LEGEND =====
chart_sheet3 = workbook.add_worksheet('Charts_Q1-Q6')
chart_sheet4 = workbook.add_worksheet('Charts_Q7-Q12')
chart_sheet5 = workbook.add_worksheet('Charts_Q13-Q18')

chart_sheets = {
    range(1, 7): chart_sheet3,
    range(7, 13): chart_sheet4,
    range(13, 19): chart_sheet5
}

# Response labels for creating legend series
response_legend = {
    1: '1 = Strongly Disagree',
    2: '2 = Disagree',
    3: '3 = Neutral',
    4: '4 = Agree',
    5: '5 = Strongly Agree',
    6: '6 = Not applicable'
}

for q_range, sheet in chart_sheets.items():
    data_row = 2
    
    for q_num in q_range:
        q_data = usability_summary[usability_summary['Question_Number'] == f'Q{q_num}']
        
        if not q_data.empty:
            # Sort by response order
            response_order = ['Strongly Disagree (1)', 'Disagree (2)', 'Neutral (3)', 
                             'Agree (4)', 'Strongly Agree (5)', 'Not applicable', 'Not applicable ']
            q_data_sorted = q_data.copy()
            q_data_sorted['Response'] = pd.Categorical(q_data_sorted['Response'], 
                                                       categories=response_order, 
                                                       ordered=True)
            q_data_sorted = q_data_sorted.sort_values('Response')
            q_data_sorted = q_data_sorted.reset_index(drop=True)
            
            # Map responses to numbers for X-axis
            response_to_number = {
                'Strongly Disagree (1)': 1,
                'Disagree (2)': 2,
                'Neutral (3)': 3,
                'Agree (4)': 4,
                'Strongly Agree (5)': 5,
                'Not applicable': 6,
                'Not applicable ': 6
            }
            
            # Write data with legend labels
            start_row = data_row
            sheet.write_row(f'B{start_row}', ['Response_Number', 'Legend_Label', 'Count'])
            for i, row in enumerate(q_data_sorted.itertuples(), start=start_row+1):
                resp_num = response_to_number.get(row.Response, 0)
                legend_label = response_legend.get(resp_num, str(resp_num))
                sheet.write_row(f'B{i}', [resp_num, legend_label, row.Count])
            
            # Create chart
            chart = workbook.add_chart({'type': 'column'})
            
            # Add series for each response category with proper legend
            for i, row in enumerate(q_data_sorted.itertuples()):
                resp_num = response_to_number.get(row.Response, 0)
                legend_label = response_legend.get(resp_num, str(resp_num))
                
                chart.add_series({
                    'name': legend_label,
                    'categories': f'={sheet.name}!$B${start_row+1+i}:$B${start_row+1+i}',
                    'values': f'={sheet.name}!$D${start_row+1+i}:$D${start_row+1+i}',
                    'data_labels': {'value': True, 'font': {'name': 'CMU Serif', 'size': 9}},
                })
            
            # Get full question text
            question_text = question_texts.get(f'Q{q_num}', f'Question {q_num}')
            
            # Set chart formatting
            chart.set_title({
                'name': question_text,
                'name_font': {'name': 'CMU Serif', 'size': 10}
            })
            chart.set_x_axis({
                'name': 'Response',
                'name_font': {'name': 'CMU Serif', 'size': 10},
                'num_font': {'name': 'CMU Serif', 'size': 9}
            })
            chart.set_y_axis({
                'name': 'Number of Participants',
                'name_font': {'name': 'CMU Serif', 'size': 10},
                'num_font': {'name': 'CMU Serif', 'size': 9}
            })
            chart.set_legend({
                'position': 'bottom',
                'font': {'name': 'CMU Serif', 'size': 8}
            })
            chart.set_size({'width': 550, 'height': 450})
            
            # Position charts in grid (2 columns, 3 rows per sheet)
            sheet_q_num = q_num - list(q_range)[0]
            col_offset = sheet_q_num % 2
            row_offset = (sheet_q_num // 2) * 28
            
            col_letter = chr(66 + col_offset * 10)
            sheet.insert_chart(f'{col_letter}{2 + row_offset}', chart)
            
            data_row = start_row + len(q_data_sorted) + 3

print(f"  ✓ Individual question charts (Q1-Q18, across 3 sheets)")

# Close the Pandas Excel writer and output the Excel file
writer.close()

print("\n" + "="*70)
print("✓ COMPLETE!")
print("="*70)
print(f"\n📁 File created: {output_file}")
print(f"\n📊 Sheets included:")
print("   1. Demographics - Raw data")
print("   2. Usability - Raw data")
print("   3. Demo_Summary - Categorical counts (with full question text)")
print("   4. Usability_Summary - Response distributions (with full question text)")
print("   5. Usability_Averages - Average scores (with full question text)")
print("   6. Charts_Demographics - Country, Gender, GenAI frequency")
print("   7. Charts_Usability - Average scores, Top/Bottom 5")
print("   8. Charts_Q1-Q6 - Individual question distributions")
print("   9. Charts_Q7-Q12 - Individual question distributions")
print("  10. Charts_Q13-Q18 - Individual question distributions")
print("\n✨ NEW FORMATTING:")
print("   ✓ Legend INSIDE each chart (not on sheet)")
print("   ✓ Title font: CMU Serif, size 10")
print("   ✓ All chart fonts: CMU Serif")
print("   ✓ Legend at bottom of chart showing response meanings")
print(f"\n📊 Total: 21 charts created")
