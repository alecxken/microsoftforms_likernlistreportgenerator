# Solutions Evaluation Report Generator

A Flask application for generating interactive evaluation reports from survey data. This tool allows users to upload evaluation data, define custom sections and questions, and generate comprehensive interactive reports to analyze and compare different solutions.

## Features

- Upload Excel/CSV data with evaluation scores
- Define custom evaluation sections and weights
- Define custom evaluation questions
- Generate interactive reports with charts and analysis
- Adjust weightings to see how rankings change
- Detailed analysis by section, question, and participant
- Export templates to standardize data collection

## Getting Started

### Prerequisites

- Python 3.7+
- Flask
- Pandas
- Werkzeug

### Installation

1. Clone the repository:

```bash
git clone https://github.com/yourusername/solutions-evaluation-report-generator.git
cd solutions-evaluation-report-generator
```

2. Create a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install flask pandas werkzeug
```

4. Set up the template files:

```bash
python setup_templates.py
```

5. Run the application:

```bash
python app.py
```

6. Open your browser and navigate to `http://localhost:5000`

## Usage

### Step 1: Define Evaluation Structure

First, define the structure of your evaluation by uploading:
- **Sections file**: Defines the sections of your evaluation and their default weights
- **Questions file**: Defines the questions in each section

You can download template files from the application to get started.

### Step 2: Upload Evaluation Data

Upload your evaluation data in Excel format. The data should include:
- Participant information
- Scores for each vendor/solution for each question
- Any standalone questions (comments, preferences, etc.)

### Step 3: View the Report

Once your data is processed, you can view the interactive report. The report includes:
- Executive summary with key findings
- Dynamic comparison with adjustable section weights
- Detailed analysis by section, question, and participant
- Insights and methodology

## Data Format

### Sections Template

| section_id | title | default_weight |
|------------|-------|----------------|
| A | Evaluation Based User Needs | 40 |
| B | User Interface & User Experience | 30 |
| C | Technical Architecture & Integration | 30 |

### Questions Template

| id | title | section | description |
|----|-------|---------|-------------|
| 1 | Payment Capabilities | A | How well does the solution handle payments? |
| 2 | User Interface | B | How intuitive is the interface? |
| 3 | API Integration | C | How comprehensive is the API ecosystem? |

### Data Template

Your Excel file should contain:
- Columns for participant info (Name, Email, etc.)
- Score columns for each vendor and question
- Standalone question columns (Comments, Preferences, etc.)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Bootstrap for the UI components
- Highcharts for the charts and visualizations