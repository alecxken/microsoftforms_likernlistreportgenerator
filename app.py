from flask import Flask, request, jsonify, render_template, send_from_directory
import pandas as pd
import json
import re
import os
from werkzeug.utils import secure_filename
from flask import send_file
import matplotlib
matplotlib.use('Agg')  # Set the backend to Agg for server environments
import matplotlib.pyplot as plt
import numpy as np
import tempfile
import base64
from io import BytesIO
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from datetime import datetime

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
        participant_types_file = request.files.get('participant_types_file')  # New participant types file
        
        sections = {}
        questions = []
        participant_types = {}  # Dictionary to store participant type information
        
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
                
                if pd.isna(description):
                    description = ""  # Replace NaN with empty string
                
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
        
        # Load participant types if provided
        if participant_types_file and allowed_file(participant_types_file.filename):
            participant_types_filename = secure_filename(participant_types_file.filename)
            participant_types_filepath = os.path.join(app.config['UPLOAD_FOLDER'], participant_types_filename)
            participant_types_file.save(participant_types_filepath)
            
            if participant_types_filepath.endswith('.csv'):
                participant_types_df = pd.read_csv(participant_types_filepath)
            else:
                participant_types_df = pd.read_excel(participant_types_filepath)
            
            for _, row in participant_types_df.iterrows():
                name = row.get('Name', '')
                participant_type = row.get('Type', '')
                
                if name and participant_type:
                    participant_types[name] = participant_type
        
        # Extract data from Excel
        try:
            result = extract_data(filepath, standalone_questions, sections, questions, participant_types)
            
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

def extract_data(excel_path, standalone_questions, sections_data={}, questions_data=[], participant_types={}):
    """Extract evaluation data from Excel file with participant type support"""
    df = pd.read_excel(excel_path)

    # Replace NaN with empty strings in the entire dataframe
    df = df.fillna("")

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
        if not value:  # Empty string or None
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

    # Extracting Question Scores - Initialize as dictionaries containing lists
    question_scores = {str(q["id"]): {v: [] for v in vendors} for q in questions}

    for _, row in df.iterrows():
        for q in [question["id"] for question in questions]:
            for vendor in vendors:
                column_name = vendor if q == 1 else f"{vendor}{q}"
                if column_name in row:
                    score = extract_score(row[column_name])
                    if score is not None:
                        # Append to the list instead of assigning
                        question_scores[str(q)][vendor].append(score)

    # Calculating average scores per question
    question_averages = {}
    for q in question_scores:
        question_averages[q] = {
            vendor: round(sum(scores) / len(scores), 2) if scores else 0
            for vendor, scores in question_scores[q].items()
        }

    # Extract participant scores
    participants = df["Name"].unique().tolist()
    participant_scores = {name: {v: 0 for v in vendors} for name in participants if name}

    # Track participants by type
    core_participants = []
    noncore_participants = []
    
    for _, row in df.iterrows():
        name = row.get("Name")
        if name and name in participant_scores:
            # Determine participant type
            participant_type = participant_types.get(name, "Unknown")
            
            # Add to appropriate list
            if participant_type == "Core":
                if name not in core_participants:
                    core_participants.append(name)
            elif participant_type == "Non-core":
                if name not in noncore_participants:
                    noncore_participants.append(name)
            
            # Calculate scores
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
            if question in row and row[question]:
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

    # Calculate Core vs Non-core comparison data
    core_vs_noncore_analysis = calculate_core_vs_noncore_analysis(
        df, vendors, questions, participant_types, core_participants, noncore_participants
    )

    # Calculate participant question and section scores for enhanced individual analysis
    participant_question_scores = {}
    participant_section_scores = {}
    
    # Process question scores by participant
    for participant in participants:
        if not participant:
            continue
            
        participant_question_scores[participant] = {}
        
        for q in questions:
            q_id = str(q["id"])
            participant_question_scores[participant][q_id] = {}
            
            for vendor in vendors:
                scores = []
                for _, row in df.iterrows():
                    if row.get("Name") == participant:
                        column_name = vendor if q["id"] == 1 else f"{vendor}{q['id']}"
                        if column_name in row:
                            score = extract_score(row[column_name])
                            if score is not None:
                                scores.append(score)
                                break  # Each participant should only have one row
                
                if scores:
                    participant_question_scores[participant][q_id][vendor] = scores[0]
                else:
                    participant_question_scores[participant][q_id][vendor] = 0
    
    # Calculate section scores by participant
    for participant in participants:
        if not participant:
            continue
            
        participant_section_scores[participant] = {}
        
        for section_key, section in sections_data.items():
            participant_section_scores[participant][section_key] = {}
            
            for vendor in vendors:
                question_scores_sum = 0
                question_count = 0
                
                for q_id in section.get("questions", []):
                    str_q_id = str(q_id)
                    if str_q_id in participant_question_scores[participant] and vendor in participant_question_scores[participant][str_q_id]:
                        q_score = participant_question_scores[participant][str_q_id][vendor]
                        if q_score > 0:  # Only count non-zero scores
                            question_scores_sum += q_score
                            question_count += 1
                
                if question_count > 0:
                    participant_section_scores[participant][section_key][vendor] = round(question_scores_sum / question_count, 2)
                else:
                    participant_section_scores[participant][section_key][vendor] = 0

    # Compile JSON Output
    output_data = {
        "vendors": vendors,
        "participants": participants,
        "questions": questions,
        "question_scores": question_averages,
        "participant_scores": participant_scores,
        "standalone_questions": standalone_responses,
        "sections": sections_data,
        "participant_types": {name: participant_types.get(name, "Unknown") for name in participants if name},
        "core_participants": core_participants,
        "noncore_participants": noncore_participants,
        "core_vs_noncore_analysis": core_vs_noncore_analysis,
        "participant_question_scores": participant_question_scores,
        "participant_section_scores": participant_section_scores
    }

    return output_data

def calculate_core_vs_noncore_analysis(df, vendors, questions, participant_types, core_participants, noncore_participants):
    """Calculate analysis comparing Core vs Non-core participant responses"""
    
    # Initialize analysis structure
    analysis = {
        "overall_scores": {
            "Core": {v: [] for v in vendors},  # Changed to initialize as empty lists
            "Non-core": {v: [] for v in vendors}
        },
        "question_scores": {
            str(q["id"]): {
                "Core": {v: [] for v in vendors},  # Changed to initialize as empty lists
                "Non-core": {v: [] for v in vendors}
            } for q in questions
        },
        "preference_distribution": {
            "Core": {v: 0 for v in vendors},
            "Non-core": {v: 0 for v in vendors}
        }
    }
    
    # Helper function to extract numeric score
    def extract_score(value):
        if not value:  # Empty string or None
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
    
    # Process each participant's scores
    for _, row in df.iterrows():
        name = row.get("Name")
        if not name:
            continue
            
        # Determine participant type
        participant_type = participant_types.get(name, "Unknown")
        if participant_type not in ["Core", "Non-core"]:
            continue
        
        # Calculate vendor scores for this participant
        vendor_scores = {v: [] for v in vendors}
        
        for q in [question["id"] for question in questions]:
            for vendor in vendors:
                column_name = vendor if q == 1 else f"{vendor}{q}"
                if column_name in row:
                    score = extract_score(row[column_name])
                    if score is not None:
                        vendor_scores[vendor].append(score)
                        
                        # Add to question-specific scores - FIXED: append to list instead of overwriting
                        analysis["question_scores"][str(q)][participant_type][vendor].append(score)
        
        # Calculate average score per vendor for this participant
        for vendor, scores in vendor_scores.items():
            if scores:
                avg_score = sum(scores) / len(scores)
                
                # Add to overall scores - FIXED: append to list instead of overwriting
                analysis["overall_scores"][participant_type][vendor].append(avg_score)
        
        # Identify preferred vendor (highest average score)
        if vendor_scores:
            avg_scores = {v: sum(scores) / len(scores) if scores else 0 for v, scores in vendor_scores.items()}
            if avg_scores:
                preferred_vendor = max(avg_scores.items(), key=lambda x: x[1])[0]
                analysis["preference_distribution"][participant_type][preferred_vendor] += 1
    
    # Calculate average overall scores
    for p_type in ["Core", "Non-core"]:
        for vendor in vendors:
            scores = analysis["overall_scores"][p_type][vendor]  # Now this is a list
            analysis["overall_scores"][p_type][vendor] = round(sum(scores) / len(scores), 2) if scores else 0
    
    # Calculate average question scores
    for q in analysis["question_scores"]:
        for p_type in ["Core", "Non-core"]:
            for vendor in vendors:
                scores = analysis["question_scores"][q][p_type][vendor]  # Now this is a list
                analysis["question_scores"][q][p_type][vendor] = round(sum(scores) / len(scores), 2) if scores else 0
    
    # Calculate top vendor by type
    analysis["top_vendor"] = {
        "Core": max(analysis["overall_scores"]["Core"].items(), key=lambda x: x[1])[0] if analysis["overall_scores"]["Core"] else None,
        "Non-core": max(analysis["overall_scores"]["Non-core"].items(), key=lambda x: x[1])[0] if analysis["overall_scores"]["Non-core"] else None
    }
    
    # Calculate agreement level (how often Core and Non-core agree on the best solution)
    core_rankings = sorted(vendors, key=lambda v: analysis["overall_scores"]["Core"].get(v, 0), reverse=True)
    noncore_rankings = sorted(vendors, key=lambda v: analysis["overall_scores"]["Non-core"].get(v, 0), reverse=True)
    
    # Kendall's Tau-like metric for agreement (simplified)
    agreement_score = 0
    max_score = len(vendors)
    
    for i, vendor in enumerate(core_rankings):
        noncore_position = noncore_rankings.index(vendor)
        agreement_score += 1 - abs(i - noncore_position) / max_score
    
    analysis["agreement_level"] = round(agreement_score / len(vendors), 2) if vendors else 0
    
    return analysis

try:
    pdfmetrics.registerFont(TTFont('Avenir-Book', 'Avenir-Book.ttf'))
    font_name = 'Avenir-Book'
except:
    font_name = 'Helvetica'

# Add this route to your Flask app
@app.route('/export-pdf')
def export_pdf():
    data_url = request.args.get('dataUrl')
    if not data_url:
        return "No data URL provided", 400
    
    # Load the JSON data
    try:
        with open(os.path.join(app.root_path, data_url.lstrip('/')), 'r') as f:
            report_data = json.load(f)
    except Exception as e:
        return f"Error loading data: {str(e)}", 500
    
    # Create a PDF
    pdf_path = generate_pdf_report(report_data)
    
    # Get filename from the data_url
    filename = os.path.splitext(os.path.basename(data_url))[0]
    
    # Send the PDF file
    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=f"{filename}_report.pdf",
        mimetype='application/pdf'
    )

def generate_pdf_report(data):
    """Generate a PDF report from the evaluation data"""
    
    # Extract data
    vendors = data.get('vendors', [])
    participants = data.get('participants', [])
    questions = data.get('questions', [])
    question_scores = data.get('question_scores', {})
    participant_scores = data.get('participant_scores', {})
    sections = data.get('sections', {})
    participant_types = data.get('participant_types', {})
    core_vs_noncore = data.get('core_vs_noncore_analysis', {})
    
    # Create a temporary file
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    
    # Create PDF document
    doc = SimpleDocTemplate(
        temp_pdf.name,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    # Create styles
    styles = getSampleStyleSheet()
    
    # Create custom styles with Avenir font
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Title'],
        fontName=font_name,
        fontSize=16,
        spaceAfter=16
    )
    
    heading_style = ParagraphStyle(
        'Heading1',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=14,
        spaceAfter=12
    )
    
    subheading_style = ParagraphStyle(
        'Heading2',
        parent=styles['Heading2'],
        fontName=font_name,
        fontSize=12,
        spaceAfter=10
    )
    
    normal_style = ParagraphStyle(
        'Normal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=10,
        spaceAfter=8
    )
    
    # Initialize the elements list for our PDF
    elements = []
    
    # Title
    elements.append(Paragraph(f"Solutions Evaluation Report: {' vs '.join(vendors)}", title_style))
    elements.append(Spacer(1, 12))
    
    # Date and summary
    date_string = datetime.now().strftime("%B %d, %Y")
    elements.append(Paragraph(f"Generated on: {date_string}", normal_style))
    elements.append(Spacer(1, 12))
    
    # Count core and non-core participants
    core_count = sum(1 for p_type in participant_types.values() if p_type == "Core")
    noncore_count = sum(1 for p_type in participant_types.values() if p_type == "Non-core")
    
    summary_text = f"This report presents an analysis of the evaluation responses from {len(participants)} participants ({core_count} Core and {noncore_count} Non-core) who assessed {len(vendors)} solutions: {', '.join(vendors)}. The solutions were evaluated across {len(questions)} criteria organized in {len(sections)} sections."
    elements.append(Paragraph(summary_text, normal_style))
    elements.append(Spacer(1, 24))
    
    # Calculate overall scores
    section_weights = {section: sections[section].get('defaultWeight', 0) for section in sections}
    total_weight = sum(section_weights.values())
    if total_weight > 0:
        for section in section_weights:
            section_weights[section] = section_weights[section] / total_weight * 100
    
    # Calculate section scores
    section_scores = {}
    for section_id, section_info in sections.items():
        section_scores[section_id] = {}
        for vendor in vendors:
            total_score = 0
            count = 0
            for q_id in section_info.get('questions', []):
                if str(q_id) in question_scores and vendor in question_scores[str(q_id)]:
                    total_score += question_scores[str(q_id)][vendor]
                    count += 1
            section_scores[section_id][vendor] = total_score / count if count > 0 else 0
    
    # Calculate weighted scores
    weighted_scores = {}
    for vendor in vendors:
        score = 0
        for section_id, weight in section_weights.items():
            if section_id in section_scores and vendor in section_scores[section_id]:
                score += section_scores[section_id][vendor] * (weight / 100)
        weighted_scores[vendor] = score
    
    # Sort vendors by weighted score
    sorted_vendors = sorted(vendors, key=lambda v: weighted_scores.get(v, 0), reverse=True)
    
    # Executive Summary section
    elements.append(Paragraph("1. Executive Summary", heading_style))
    elements.append(Spacer(1, 8))
    
    # Key findings
    elements.append(Paragraph("Key Findings:", subheading_style))
    
    if len(sorted_vendors) > 0:
        top_vendor = sorted_vendors[0]
        top_score = weighted_scores.get(top_vendor, 0)
        elements.append(Paragraph(f"• Top solution: {top_vendor} with a weighted score of {top_score:.2f}.", normal_style))
    
    if len(sorted_vendors) > 1:
        second_vendor = sorted_vendors[1]
        second_score = weighted_scores.get(second_vendor, 0)
        elements.append(Paragraph(f"• Runner-up: {second_vendor} with a weighted score of {second_score:.2f}.", normal_style))
    
    # Add Core vs Non-core findings
    if core_vs_noncore and 'top_vendor' in core_vs_noncore:
        top_core = core_vs_noncore['top_vendor'].get('Core')
        top_noncore = core_vs_noncore['top_vendor'].get('Non-core')
        
        if top_core:
            core_score = core_vs_noncore['overall_scores']['Core'].get(top_core, 0)
            elements.append(Paragraph(f"• Core participants preferred: {top_core} with a score of {core_score:.2f}.", normal_style))
        
        if top_noncore:
            noncore_score = core_vs_noncore['overall_scores']['Non-core'].get(top_noncore, 0)
            elements.append(Paragraph(f"• Non-core participants preferred: {top_noncore} with a score of {noncore_score:.2f}.", normal_style))
        
        if 'agreement_level' in core_vs_noncore:
            agreement = core_vs_noncore['agreement_level']
            agreement_text = "High" if agreement > 0.8 else "Moderate" if agreement > 0.5 else "Low"
            elements.append(Paragraph(f"• Agreement level between Core and Non-core participants: {agreement_text} ({agreement:.2f}).", normal_style))
    
    elements.append(Spacer(1, 12))
    
    # Add chart for overall scores
    chart_path = create_bar_chart(
        sorted_vendors, 
        [weighted_scores.get(v, 0) for v in sorted_vendors],
        "Overall Weighted Scores",
        "Solution",
        "Score (1-5)"
    )
    elements.append(Paragraph("Overall Solution Rankings:", subheading_style))
    elements.append(Image(chart_path, width=400, height=300))
    elements.append(Spacer(1, 16))
    
    # Add Core vs Non-core comparison chart
    if core_vs_noncore and 'overall_scores' in core_vs_noncore:
        core_scores = [core_vs_noncore['overall_scores']['Core'].get(v, 0) for v in sorted_vendors]
        noncore_scores = [core_vs_noncore['overall_scores']['Non-core'].get(v, 0) for v in sorted_vendors]
        
        comparison_path = create_group_bar_chart(
            sorted_vendors,
            [core_scores, noncore_scores],
            ["Core", "Non-core"],
            "Core vs Non-core Evaluation",
            "Solution",
            "Score (1-5)"
        )
        elements.append(Paragraph("Core vs Non-core Comparison:", subheading_style))
        elements.append(Image(comparison_path, width=400, height=300))
     # Continuing from where the previous part left off

    # Section Performance
    elements.append(Paragraph("2. Section Performance", heading_style))
    elements.append(Spacer(1, 8))
    
    for section_id, section_info in sections.items():
        section_title = section_info.get('title', f'Section {section_id}')
        elements.append(Paragraph(f"Section {section_id}: {section_title}", subheading_style))
        
        # Create a table for section scores
        section_data = [["Solution", "Score"]]
        section_vendors = sorted(
            vendors, 
            key=lambda v: section_scores.get(section_id, {}).get(v, 0), 
            reverse=True
        )
        
        for vendor in section_vendors:
            score = section_scores.get(section_id, {}).get(vendor, 0)
            section_data.append([vendor, f"{score:.2f}"])
        
        # Create the table
        table = Table(section_data, colWidths=[300, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 8))
        
        # Create a chart for the section
        chart_path = create_bar_chart(
            section_vendors, 
            [section_scores.get(section_id, {}).get(v, 0) for v in section_vendors],
            f"{section_title} Scores",
            "Solution",
            "Score (1-5)"
        )
        elements.append(Image(chart_path, width=400, height=200))
        elements.append(Spacer(1, 16))
    
    # Question Performance
    elements.append(Paragraph("3. Question Performance", heading_style))
    elements.append(Spacer(1, 8))
    
    # Group questions by section
    questions_by_section = {}
    for q in questions:
        section = q.get('section')
        if section not in questions_by_section:
            questions_by_section[section] = []
        questions_by_section[section].append(q)
    
    for section_id, section_questions in questions_by_section.items():
        section_title = sections.get(section_id, {}).get('title', f'Section {section_id}')
        elements.append(Paragraph(f"Section {section_id}: {section_title}", subheading_style))
        elements.append(Spacer(1, 8))
        
        for question in section_questions:
            q_id = str(question.get('id'))
            q_title = question.get('title', f'Question {q_id}')
            
            elements.append(Paragraph(f"Q{q_id}: {q_title}", normal_style))
            
            if q_id in question_scores:
                # Sort vendors by question score
                q_vendors = sorted(
                    vendors, 
                    key=lambda v: question_scores[q_id].get(v, 0), 
                    reverse=True
                )
                
                # Create a table for question scores
                q_data = [["Solution", "Score"]]
                for vendor in q_vendors:
                    score = question_scores[q_id].get(vendor, 0)
                    q_data.append([vendor, f"{score:.2f}"])
                
                # Create the table
                table = Table(q_data, colWidths=[300, 100])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), font_name),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 8))
            
            elements.append(Spacer(1, 8))
        
        elements.append(Spacer(1, 16))
    
    # Core vs Non-core Analysis Section (New)
    if core_vs_noncore:
        elements.append(Paragraph("4. Core vs Non-core Analysis", heading_style))
        elements.append(Spacer(1, 8))
        
        # Add explanation
        elements.append(Paragraph("This section compares the evaluations from Core and Non-core participants to identify any differences in preferences or priorities.", normal_style))
        elements.append(Spacer(1, 12))
        
        # Overall scores comparison
        elements.append(Paragraph("Overall Scores Comparison:", subheading_style))
        
        # Create comparison table
        comparison_data = [["Solution", "Core Score", "Non-core Score", "Difference"]]
        
        for vendor in sorted_vendors:
            core_score = core_vs_noncore['overall_scores']['Core'].get(vendor, 0)
            noncore_score = core_vs_noncore['overall_scores']['Non-core'].get(vendor, 0)
            difference = core_score - noncore_score
            
            comparison_data.append([
                vendor, 
                f"{core_score:.2f}", 
                f"{noncore_score:.2f}", 
                f"{difference:+.2f}"
            ])
        
        # Create the table
        table = Table(comparison_data, colWidths=[200, 100, 100, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (3, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 12))
        
        # Preference distribution
        elements.append(Paragraph("Preference Distribution:", subheading_style))
        
        # Create preference chart
        pref_chart_path = create_preference_chart(
            vendors,
            [core_vs_noncore['preference_distribution']['Core'].get(v, 0) for v in vendors],
            [core_vs_noncore['preference_distribution']['Non-core'].get(v, 0) for v in vendors],
            "Top Choice Distribution",
            "Solution",
            "Number of Participants"
        )
        elements.append(Image(pref_chart_path, width=400, height=300))
        elements.append(Spacer(1, 12))
        
        # Key differences in questions
        elements.append(Paragraph("Key Differences in Question Scores:", subheading_style))
        
        # Find questions with significant differences
        significant_questions = []
        
        for q in questions:
            q_id = str(q.get('id'))
            
            if q_id in core_vs_noncore['question_scores']:
                max_diff = 0
                
                for vendor in vendors:
                    core_score = core_vs_noncore['question_scores'][q_id]['Core'].get(vendor, 0)
                    noncore_score = core_vs_noncore['question_scores'][q_id]['Non-core'].get(vendor, 0)
                    diff = abs(core_score - noncore_score)
                    
                    if diff > max_diff:
                        max_diff = diff
                
                if max_diff >= 1.0:  # Consider difference of 1.0 or more as significant
                    significant_questions.append((q, max_diff))
        
        # Sort by difference (largest first)
        significant_questions.sort(key=lambda x: x[1], reverse=True)
        
        # Show top 5 questions with differences
        if significant_questions:
            # Table header
            diff_data = [["Question", "Solution", "Core", "Non-core", "Diff"]]
            
            for q, diff in significant_questions[:5]:
                q_id = str(q.get('id'))
                q_title = q.get('title', '')
                
                # Find vendor with the maximum difference
                max_diff_vendor = None
                max_diff_value = 0
                
                for vendor in vendors:
                    core_score = core_vs_noncore['question_scores'][q_id]['Core'].get(vendor, 0)
                    noncore_score = core_vs_noncore['question_scores'][q_id]['Non-core'].get(vendor, 0)
                    diff = abs(core_score - noncore_score)
                    
                    if diff > max_diff_value:
                        max_diff_value = diff
                        max_diff_vendor = vendor
                
                if max_diff_vendor:
                    core_score = core_vs_noncore['question_scores'][q_id]['Core'].get(max_diff_vendor, 0)
                    noncore_score = core_vs_noncore['question_scores'][q_id]['Non-core'].get(max_diff_vendor, 0)
                    diff = core_score - noncore_score
                    
                    diff_data.append([
                        f"Q{q_id}: {q_title[:30]}...",
                        max_diff_vendor,
                        f"{core_score:.2f}",
                        f"{noncore_score:.2f}",
                        f"{diff:+.2f}"
                    ])
            
            # Create the table
            table = Table(diff_data, colWidths=[200, 100, 70, 70, 60])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (2, 0), (4, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), font_name),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ]))
            elements.append(table)
        else:
            elements.append(Paragraph("No significant differences found in question scores between Core and Non-core participants.", normal_style))
        
        elements.append(Spacer(1, 16))
    
    # Participant Assessments
    elements.append(Paragraph("5. Participant Assessments", heading_style))
    elements.append(Spacer(1, 8))
    
    # Average scores by participant
    participant_avgs = {}
    for p_name, scores in participant_scores.items():
        participant_avgs[p_name] = {
            vendor: scores.get(vendor, 0)
            for vendor in vendors
        }
    
    # Create a heatmap for participant scores
    if len(participants) > 0 and len(vendors) > 0:
        # Sort participants by type for better visualization
        sorted_participants = []
        
        # First core participants
        for p in participants:
            if p and participant_types.get(p, "") == "Core":
                sorted_participants.append(p)
        
        # Then non-core participants
        for p in participants:
            if p and participant_types.get(p, "") == "Non-core":
                sorted_participants.append(p)
        
        # Then others
        for p in participants:
            if p and p not in sorted_participants:
                sorted_participants.append(p)
        
        # Create participant labels with type
        participant_labels = [f"{p} ({participant_types.get(p, 'Unknown')})" for p in sorted_participants if p]
        
        heatmap_path = create_heatmap(
            participant_labels, 
            vendors,
            [[participant_avgs.get(p, {}).get(v, 0) for v in vendors] for p in sorted_participants if p],
            "Participant Evaluation Heatmap",
            "Participant",
            "Solution"
        )
        elements.append(Image(heatmap_path, width=450, height=300))
        elements.append(Spacer(1, 16))
    
    # Calculate top choice by participant
    top_choices = {}
    for p_name, scores in participant_avgs.items():
        if scores:
            top_vendor = max(scores.items(), key=lambda x: x[1])[0]
            top_choices[p_name] = top_vendor
    
    # Count vendor selections
    vendor_counts = {v: 0 for v in vendors}
    for vendor in top_choices.values():
        if vendor in vendor_counts:
            vendor_counts[vendor] += 1
    
    # Create pie chart for top choices
    if sum(vendor_counts.values()) > 0:
        pie_path = create_pie_chart(
            vendor_counts.keys(),
            vendor_counts.values(),
            "Participants' Top Choices"
        )
        elements.append(Paragraph("Participants' Top Choices:", subheading_style))
        elements.append(Image(pie_path, width=400, height=300))
        elements.append(Spacer(1, 16))
    
    # Individual Participant Analysis (New)
    elements.append(Paragraph("6. Individual Participant Analysis", heading_style))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("This section provides insights into individual participant evaluations, showing variations in how different evaluators rated the solutions.", normal_style))
    elements.append(Spacer(1, 12))
    
    # Select a few representative participants for the PDF report
    selected_participants = []
    
    # Try to include at least one Core and one Non-core participant
    core_selected = False
    noncore_selected = False
    
    for p in participants:
        if participant_types.get(p, "") == "Core" and not core_selected:
            selected_participants.append(p)
            core_selected = True
        elif participant_types.get(p, "") == "Non-core" and not noncore_selected:
            selected_participants.append(p)
            noncore_selected = True
            
        if len(selected_participants) >= 2:
            break
    
    # Add more if needed
    if len(selected_participants) < 2 and len(participants) > 2:
        for p in participants:
            if p not in selected_participants:
                selected_participants.append(p)
                if len(selected_participants) >= 2:
                    break
    
    # For each selected participant, include a small analysis
    for participant in selected_participants:
        participant_type = participant_types.get(participant, "Unknown")
        
        elements.append(Paragraph(f"Participant: {participant} ({participant_type})", subheading_style))
        
        # Create a bar chart showing this participant's ratings compared to group average
        comparison_data = []
        participant_data = []
        group_data = []
        
        for vendor in vendors:
            participant_score = participant_scores.get(participant, {}).get(vendor, 0)
            participant_data.append(participant_score)
            
            # Calculate group average
            group_scores = []
            for p in participants:
                if p != participant:  # Exclude this participant
                    score = participant_scores.get(p, {}).get(vendor, 0)
                    if score > 0:
                        group_scores.append(score)
            
            group_avg = sum(group_scores) / len(group_scores) if group_scores else 0
            group_data.append(group_avg)
        
        # Create comparison chart
        chart_path = create_group_bar_chart(
            vendors,
            [participant_data, group_data],
            ["Participant", "Group Average"],
            f"{participant}'s Evaluation vs Group Average",
            "Solution",
            "Score (1-5)"
        )
        elements.append(Image(chart_path, width=400, height=250))
        elements.append(Spacer(1, 12))
        
        # Add overall insight
        top_vendor = max(participant_scores.get(participant, {}).items(), key=lambda x: x[1])[0] if participant_scores.get(participant, {}) else "N/A"
        top_score = participant_scores.get(participant, {}).get(top_vendor, 0)
        
        elements.append(Paragraph(f"Top choice: {top_vendor} (Score: {top_score:.2f})", normal_style))
        elements.append(Spacer(1, 16))
    
    # Methodology
    elements.append(Paragraph("7. Evaluation Methodology", heading_style))
    elements.append(Spacer(1, 8))
    
    methodology_text = f"""
    This evaluation was conducted with {len(participants)} participants who assessed {len(vendors)} solutions: {', '.join(vendors)}.
    The participants included {core_count} Core users and {noncore_count} Non-core users.
    The assessment was structured across {len(sections)} sections with a total of {len(questions)} evaluation criteria.
    
    Participants rated each solution on a scale of 1-5:
    1 - Poor: Fails to meet basic requirements
    2 - Basic: Meets minimal requirements with limited functionality
    3 - Standard: Satisfactory implementation with adequate functionality
    4 - Good: Strong implementation with comprehensive functionality
    5 - Excellent: Outstanding implementation with superior functionality
    
    Section weights used for the overall rankings:
    """
    elements.append(Paragraph(methodology_text, normal_style))
    
    # Section weights table
    weights_data = [["Section", "Title", "Weight"]]
    for section_id, section_info in sections.items():
        title = section_info.get('title', f'Section {section_id}')
        weight = section_weights.get(section_id, 0)
        weights_data.append([f"Section {section_id}", title, f"{weight:.0f}%"])
    
    # Create the table
    table = Table(weights_data, colWidths=[100, 300, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
    ]))
    elements.append(table)
    
    # Build the PDF
    doc.build(elements)
    
    return temp_pdf.name

def create_bar_chart(categories, values, title, xlabel, ylabel):
    """Create a bar chart and return the path to the saved image"""
    plt.figure(figsize=(8, 6))
    
    # Create bar plot
    bars = plt.bar(categories, values)
    
    # Add value labels on top of bars
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                f'{height:.2f}', ha='center', fontsize=9)
    
    # Add labels and title
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.ylim(0, 5.5)  # Assuming 5-point scale with space for labels
    
    # Rotate x-axis labels for better readability if needed
    if len(categories) > 3:
        plt.xticks(rotation=45, ha='right')
    
    plt.tight_layout()
    
    # Save the chart to a temporary file
    temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    plt.savefig(temp_chart.name, dpi=150)
    plt.close()
    
    return temp_chart.name

def create_group_bar_chart(categories, values_sets, group_labels, title, xlabel, ylabel):
    """Create a grouped bar chart comparing multiple datasets"""
    plt.figure(figsize=(10, 6))
    
    # Number of groups and space for bars
    n_groups = len(categories)
    bar_width = 0.35
    
    # Set positions of bars on X axis
    positions = np.arange(n_groups)
    
    # Create bars
    for i, values in enumerate(values_sets):
        offset = (i - len(values_sets)/2 + 0.5) * bar_width
        bars = plt.bar(positions + offset, values, bar_width, 
                       label=group_labels[i])
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{height:.2f}', ha='center', fontsize=8)
    
    # Add labels and title
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.ylim(0, 5.5)  # Assuming 5-point scale with space for labels
    
    # Set the x-ticks positions
    plt.xticks(positions, categories)
    
    # Rotate x-axis labels for better readability if needed
    if len(categories) > 3:
        plt.xticks(rotation=45, ha='right')
    
    # Add legend
    plt.legend()
    
    plt.tight_layout()
    
    # Save the chart to a temporary file
    temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    plt.savefig(temp_chart.name, dpi=150)
    plt.close()
    
    return temp_chart.name

def create_preference_chart(categories, values1, values2, title, xlabel, ylabel):
    """Create a bar chart showing preference distribution"""
    plt.figure(figsize=(10, 6))
    
    # Number of groups
    n_groups = len(categories)
    bar_width = 0.35
    
    # Set positions for bars
    index = np.arange(n_groups)
    
    # Create bars
    bars1 = plt.bar(index, values1, bar_width, label='Core')
    bars2 = plt.bar(index + bar_width, values2, bar_width, label='Non-core')
    
    # Add value labels
    for bar in bars1:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                f'{int(height)}', ha='center', fontsize=9)
    
    for bar in bars2:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                f'{int(height)}', ha='center', fontsize=9)
    
    # Add labels and title
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    
    # Set x-axis ticks
    plt.xticks(index + bar_width / 2, categories)
    
    # Rotate x-axis labels for better readability if needed
    if len(categories) > 3:
        plt.xticks(rotation=45, ha='right')
    
    # Add legend
    plt.legend()
    
    plt.tight_layout()
    
    # Save the chart to a temporary file
    temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    plt.savefig(temp_chart.name, dpi=150)
    plt.close()
    
    return temp_chart.name

def create_heatmap(rows, columns, data, title, ylabel, xlabel):
    """Create a heatmap and return the path to the saved image"""
    plt.figure(figsize=(10, max(6, len(rows) * 0.4)))  # Adjust height based on number of rows
    
    # Create heatmap
    plt.imshow(data, cmap='YlGn', aspect='auto', vmin=1, vmax=5)
    
    # Add colorbar
    cbar = plt.colorbar()
    cbar.set_label('Score (1-5)')
    
    # Add labels and title
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    
    # Set tick positions and labels
    plt.yticks(range(len(rows)), rows)
    plt.xticks(range(len(columns)), columns)
    
    # Add text annotations in each cell
    for i in range(len(rows)):
        for j in range(len(columns)):
            plt.text(j, i, f"{data[i][j]:.1f}", 
                    ha="center", va="center", color="black", fontsize=8)
    
    plt.tight_layout()
    
    # Save the chart to a temporary file
    temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    plt.savefig(temp_chart.name, dpi=150)
    plt.close()
    
    return temp_chart.name

def create_pie_chart(labels, values, title):
    """Create a pie chart and return the path to the saved image"""
    plt.figure(figsize=(8, 6))
    
    # Filter out zero values
    non_zero_data = [(label, value) for label, value in zip(labels, values) if value > 0]
    
    if non_zero_data:
        labels, values = zip(*non_zero_data)
        
        # Create pie chart
        plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, 
                shadow=False, explode=[0.05]*len(labels))
        
        # Add title
        plt.title(title)
        plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        
        plt.tight_layout()
        
        # Save the chart to a temporary file
        temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        plt.savefig(temp_chart.name, dpi=150)
        plt.close()
        
        return temp_chart.name
    else:
        # Create an empty chart if no data
        plt.text(0.5, 0.5, 'No data available', ha='center', va='center')
        
        # Save the chart to a temporary file
        temp_chart = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        plt.savefig(temp_chart.name, dpi=150)
        plt.close()
        
        return temp_chart.name

@app.template_filter('calculate_weighted_scores')
def calculate_weighted_scores(data):
    """Calculate weighted scores for all vendors and return scores with sorted vendors"""
    vendors = data.get('vendors', [])
    sections = data.get('sections', {})
    question_scores = data.get('question_scores', {})
    
    # Calculate section scores
    section_scores = {}
    for section_id, section_info in sections.items():
        section_scores[section_id] = {}
        for vendor in vendors:
            total_score = 0
            count = 0
            for q_id in section_info.get('questions', []):
                if str(q_id) in question_scores and vendor in question_scores[str(q_id)]:
                    total_score += question_scores[str(q_id)][vendor]
                    count += 1
            section_scores[section_id][vendor] = total_score / count if count > 0 else 0
    
    # Calculate weighted scores
    weighted_scores = {}
    for vendor in vendors:
        score = 0
        for section_id, section_info in sections.items():
            weight = section_info.get('defaultWeight', 0)
            if section_id in section_scores and vendor in section_scores[section_id]:
                score += section_scores[section_id][vendor] * (weight / 100)
        weighted_scores[vendor] = score
    
    # Also return the sorted vendors
    sorted_vendors = sorted(vendors, key=lambda v: weighted_scores.get(v, 0), reverse=True)
    
    return {"scores": weighted_scores, "sorted_vendors": sorted_vendors}
# Add this simple print-report route to app.py

@app.route('/print-report')
def print_report():
    """Render a print-friendly version of the report for PDF export"""
    data_url = request.args.get('dataUrl')
    if not data_url:
        return "No data URL provided", 400
    
    # Load the JSON data
    try:
        with open(os.path.join(app.root_path, data_url.lstrip('/')), 'r') as f:
            report_data = json.load(f)
    except Exception as e:
        return f"Error loading data: {str(e)}", 500
    
    # Render the print-friendly template with the report data
    return render_template('print_report.html', 
                          data=report_data, 
                          report_date=datetime.now().strftime("%B %d, %Y"))

if __name__ == '__main__':
    app.run(debug=True, port=5000)