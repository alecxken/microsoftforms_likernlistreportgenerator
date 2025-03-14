from flask import Flask, request, jsonify, render_template, send_from_directory
import pandas as pd
import json
import re
import os
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join('static', 'data'), exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

# Common non-vendor columns that should always be excluded
COMMON_NON_VENDORS = [
    'Comments', 'Language', 'Feedback', 'Suggestions', 
    'Notes', 'Remarks', 'Additional Comments', 'Further Information'
]

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

# Add the missing report-viewer route
@app.route('/report-viewer')
def report_viewer():
    return render_template('report_viewer.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/process-excel', methods=['POST'])
def process_excel():
    if 'excel_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['excel_file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Get standalone questions, supporting both comma and pipe separators
        standalone_questions_str = request.form.get('standalone_questions', '')
        if '|' in standalone_questions_str:
            standalone_questions = standalone_questions_str.split('|')
        else:
            standalone_questions = standalone_questions_str.split(',')
            
        standalone_questions = [q.strip() for q in standalone_questions if q.strip()]
        
        # Add common non-vendor columns to standalone questions
        for non_vendor in COMMON_NON_VENDORS:
            if non_vendor not in standalone_questions:
                standalone_questions.append(non_vendor)
        
        sections_file = request.files.get('sections_file')
        questions_file = request.files.get('questions_file')
        
        sections = {}
        questions = []
        
        # Load sections if provided
        if sections_file and allowed_file(sections_file.filename):
            sections_filename = secure_filename(sections_file.filename)
            sections_filepath = os.path.join(app.config['UPLOAD_FOLDER'], sections_filename)
            sections_file.save(sections_filepath)
            
            if sections_filepath.endswith('.csv'):
                sections_df = pd.read_csv(sections_filepath)
            else:
                sections_df = pd.read_excel(sections_filepath)
            
            for _, row in sections_df.iterrows():
                section_id = row.get('section_id', '')
                title = row.get('title', '')
                weight = row.get('default_weight', 33)
                
                if section_id:
                    sections[section_id] = {
                        'title': title,
                        'questions': [],
                        'defaultWeight': weight
                    }
        
        # Load questions if provided
        if questions_file and allowed_file(questions_file.filename):
            questions_filename = secure_filename(questions_file.filename)
            questions_filepath = os.path.join(app.config['UPLOAD_FOLDER'], questions_filename)
            questions_file.save(questions_filepath)
            
            if questions_filepath.endswith('.csv'):
                questions_df = pd.read_csv(questions_filepath)
            else:
                questions_df = pd.read_excel(questions_filepath)
            
            for _, row in questions_df.iterrows():
                question_id = row.get('id')
                title = row.get('title', '')
                section = row.get('section', '')
                description = row.get('description', '')
                
                if question_id and section:
                    question = {
                        'id': int(question_id),
                        'title': title,
                        'section': section,
                    }
                    
                    if description:
                        question['description'] = description
                    
                    questions.append(question)
                    
                    # Add to section questions list
                    if section in sections:
                        sections[section]['questions'].append(int(question_id))
        
        # Extract data from Excel
        try:
            result = extract_data(filepath, standalone_questions, sections, questions)
            
            # Save result to JSON file
            output_filename = f"data_{os.path.splitext(filename)[0]}.json"
            output_path = os.path.join('static', 'data', output_filename)
            
            with open(output_path, 'w') as f:
                json.dump(result, f, indent=4)
            
            return jsonify({
                'success': True,
                'message': 'Data processed successfully',
                'data_url': f'/static/data/{output_filename}'
            })
        
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file format'}), 400

def extract_data(excel_path, standalone_questions, sections_data={}, questions_data=[]):
    """Extract evaluation data from Excel file"""
    df = pd.read_excel(excel_path)

    # Extract vendor names dynamically
    base_columns = ["ID", "Start time", "Completion time", "Email", "Name", "Last modified time"]
    
    # Identify columns to exclude from vendor detection
    excluded_columns = base_columns + standalone_questions
    
    # Get all columns
    all_columns = df.columns.tolist()
    
    # First pass: identify possible numbered columns that indicate vendors
    numbered_pattern = re.compile(r'^(.*?)(\d+)$')
    vendor_with_numbers = set()
    
    for col in all_columns:
        match = numbered_pattern.match(col)
        if match and col not in excluded_columns:
            vendor_name = match.group(1).strip()
            vendor_with_numbers.add(vendor_name)
    
    # Second pass: verify vendors by checking if they have associated numbered columns
    unique_vendors = set()
    
    for vendor in vendor_with_numbers:
        # Check if this is really a vendor by looking for numbered columns
        has_numbered_columns = False
        for col in all_columns:
            if col.startswith(vendor) and re.search(r'\d+$', col):
                has_numbered_columns = True
                break
        
        if has_numbered_columns and vendor not in excluded_columns:
            unique_vendors.add(vendor)
    
    vendors = sorted(unique_vendors)  # Keep vendors sorted

    # Use provided questions or create default ones
    questions = questions_data if questions_data else [
        {"id": i, "title": f"Question {i}", "section": "A"} 
        for i in range(1, len(set([int(re.search(r'\d+$', col).group()) 
                                  for col in all_columns 
                                  if re.search(r'\d+$', col)])) + 1)
    ]

    # Helper function to extract numeric score
    def extract_score(value):
        if pd.isna(value):
            return None
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            match = re.match(r'^(\d+)', value)
            if match:
                return int(match.group(1))
            try:
                return float(value)
            except ValueError:
                return None
        return None

    # Extracting Question Scores
    question_scores = {q["id"]: {v: [] for v in vendors} for q in questions}

    for _, row in df.iterrows():
        for q in [question["id"] for question in questions]:
            for vendor in vendors:
                column_name = vendor if q == 1 else f"{vendor}{q}"
                if column_name in row:
                    score = extract_score(row[column_name])
                    if score is not None:
                        question_scores[q][vendor].append(score)

    # Calculating average scores per question
    question_averages = {}
    for q in question_scores:
        question_averages[q] = {
            vendor: round(sum(scores) / len(scores), 2) if scores else 0
            for vendor, scores in question_scores[q].items()
        }

    # Extract participant scores
    participants = df["Name"].dropna().unique().tolist()
    participant_scores = {name: {v: 0 for v in vendors} for name in participants}

    for _, row in df.iterrows():
        name = row.get("Name")
        if pd.notna(name) and name in participant_scores:
            for vendor in vendors:
                scores = []
                for q in [question["id"] for question in questions]:
                    column_name = vendor if q == 1 else f"{vendor}{q}"
                    if column_name in row:
                        score = extract_score(row[column_name])
                        if score is not None:
                            scores.append(score)
                if scores:
                    participant_scores[name][vendor] = round(sum(scores) / len(scores), 2)

    # Extract standalone question responses
    standalone_responses = {q: [] for q in standalone_questions if q in df.columns}
    for _, row in df.iterrows():
        for question in standalone_questions:
            if question in row and pd.notna(row[question]):
                if question in standalone_responses:
                    standalone_responses[question].append(row[question])

    # Use provided sections or create default ones
    if not sections_data:
        # Group questions by section
        section_groups = {}
        for q in questions:
            section = q["section"]
            if section not in section_groups:
                section_groups[section] = []
            section_groups[section].append(q["id"])
        
        # Create sections structure
        sections_data = {
            section: {
                "title": f"Section {section}",
                "questions": question_ids,
                "defaultWeight": 100 // len(section_groups)  # Equal weights
            }
            for section, question_ids in section_groups.items()
        }

    # Compile JSON Output
    output_data = {
        "vendors": vendors,
        "participants": participants,
        "questions": questions,
        "question_scores": question_averages,
        "participant_scores": participant_scores,
        "standalone_questions": standalone_responses,
        "sections": sections_data
    }

    return output_data

if __name__ == '__main__':
    app.run(debug=True, port=5000)