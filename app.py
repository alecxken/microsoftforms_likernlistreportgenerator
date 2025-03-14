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
    
    summary_text = f"This report presents an analysis of the evaluation responses from {len(participants)} participants who assessed {len(vendors)} solutions: {', '.join(vendors)}. The solutions were evaluated across {len(questions)} criteria organized in {len(sections)} sections."
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
    
    # Participant Assessments
    elements.append(Paragraph("4. Participant Assessments", heading_style))
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
        heatmap_path = create_heatmap(
            participants, 
            vendors,
            [[participant_avgs.get(p, {}).get(v, 0) for v in vendors] for p in participants],
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
    
    # Methodology
    elements.append(Paragraph("5. Evaluation Methodology", heading_style))
    elements.append(Spacer(1, 8))
    
    methodology_text = f"""
    This evaluation was conducted with {len(participants)} participants who assessed {len(vendors)} solutions: {', '.join(vendors)}.
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

def create_heatmap(rows, columns, data, title, ylabel, xlabel):
    """Create a heatmap and return the path to the saved image"""
    plt.figure(figsize=(10, 8))
    
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
if __name__ == '__main__':
    app.run(debug=True, port=5000)