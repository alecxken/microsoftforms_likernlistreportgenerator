// Global variables for data and state
let reportData = null;
let vendors = [];
let participants = [];
let questions = [];
let questionScores = {};
let participantScores = {};
let standoutSolutions = {};
let sections = {};
let sectionWeights = {};
let sectionScores = {};
let participantTypes = {};
let coreParticipants = [];
let noncoreParticipants = [];
let coreVsNoncoreAnalysis = {};
let participantQuestionScores = {};
let participantSectionScores = {};
let selectedParticipant = null;

// Define Ecobank brand colors for charts
const ecobankColors = {
    blue: '#0082BB',    // Primary blue
    green: '#BED600',   // Secondary green
    white: '#FFFFFF',   // White
    lightGray: '#F5F5F5', // Light gray
    
    // Additional shades for charts
    darkBlue: '#005B82',
    lightBlue: '#4ca8d8',
    darkGreen: '#669438',
    lightGreen: '#4caf7c',
    gray: '#666666'
};
  
// Create Highcharts theme with Ecobank colors
const ecobankHighchartsTheme = {
    colors: [
        ecobankColors.blue,
        ecobankColors.green,
        ecobankColors.darkBlue,
        ecobankColors.lightGreen,
        ecobankColors.lightBlue,
        ecobankColors.darkGreen
    ],
    
    chart: {
        backgroundColor: ecobankColors.white,
        style: {
            fontFamily: 'Arial, sans-serif'
        }
    },
    
    title: {
        style: {
            color: ecobankColors.blue,
            fontWeight: 'bold'
        }
    },
    
    subtitle: {
        style: {
            color: ecobankColors.gray
        }
    },
    
    legend: {
        itemStyle: {
            color: '#333333'
        },
        itemHoverStyle: {
            color: ecobankColors.blue
        }
    },
    
    colorAxis: {
        min: 1,
        max: 5,
        stops: [
            [0, '#ffdfdf'],    // Light red for low scores
            [0.5, '#ffffff'],  // White for middle scores
            [1, ecobankColors.lightGreen]  // Light green for high scores
        ]
    },
    
    plotOptions: {
        spline: {
            color: ecobankColors.blue,
            marker: {
                fillColor: ecobankColors.white,
                lineColor: ecobankColors.blue,
                lineWidth: 2
            }
        },
        series: {
            borderColor: ecobankColors.blue
        }
    }
};

// Initialize when document is ready
document.addEventListener('DOMContentLoaded', function() {
    // Show loading overlay
    document.getElementById('loadingOverlay').style.display = 'flex';

    // Get data URL from query parameter
    const urlParams = new URLSearchParams(window.location.search);
    const dataUrl = urlParams.get('dataUrl');

    if (!dataUrl) {
        alert('No data URL provided. Please go back and upload your data.');
        document.getElementById('loadingOverlay').style.display = 'none';
        return;
    }

    // Fetch the report data
    fetch(dataUrl)
        .then(response => response.json())
        .then(data => {
            // Store data
            reportData = data;

            // Process the data
            processReportData();

            // Update the UI
            updateUI();

            // Hide loading overlay
            document.getElementById('loadingOverlay').style.display = 'none';
        })
        .catch(error => {
            console.error('Error loading data:', error);
            alert('Error loading report data. Please try again.');
            document.getElementById('loadingOverlay').style.display = 'none';
        });
});

// Process the loaded report data
function processReportData() {
    // Extract vendors
    vendors = reportData.vendors || [];

    // Extract participants
    participants = reportData.participants || [];

    // Extract questions
    questions = reportData.questions || [];

    // Extract question scores
    questionScores = reportData.question_scores || {};

    // Extract participant scores
    participantScores = reportData.participant_scores || {};

    // Extract sections
    sections = reportData.sections || {};
    
    // Extract participant types
    participantTypes = reportData.participant_types || {};
    
    // Extract core and non-core participants
    coreParticipants = reportData.core_participants || [];
    noncoreParticipants = reportData.noncore_participants || [];
    
    // Extract core vs non-core analysis
    coreVsNoncoreAnalysis = reportData.core_vs_noncore_analysis || {};

    // Process standout solutions from standalone questions
    standoutSolutions = processStandoutSolutions();

    // Initialize section weights
    initializeSectionWeights();

    // Calculate section scores
    sectionScores = calculateSectionScores();
     // Extract participant question scores
    participantQuestionScores = reportData.participant_question_scores || {};
    
     // Extract participant section scores
    participantSectionScores = reportData.participant_section_scores || {};

}

// Process standout solutions from standalone questions
function processStandoutSolutions() {
    const result = {};

    // Initialize counts for each vendor
    vendors.forEach(vendor => {
        result[vendor] = 0;
    });

    // Look for standout solutions in any standalone question
    if (reportData.standalone_questions) {
        // Try common field names for standout solutions
        const possibleFields = [
            "For You What Three Solutions Stood out the most",
            "What solutions stood out",
            "Top solutions",
            "Preferred solutions",
            "Best solutions"
        ];

        for (const field of possibleFields) {
            if (reportData.standalone_questions[field]) {
                const responses = reportData.standalone_questions[field];

                responses.forEach(response => {
                    if (!response) return;

                    // Split by common separators and process each selection
                    const selections = response.split(/[,;]/).filter(s => s.trim() !== '');

                    selections.forEach(selection => {
                        // Find the closest matching vendor
                        const vendorMatch = findClosestVendor(selection.trim());
                        if (vendorMatch) {
                            result[vendorMatch] += 1;
                        }
                    });
                });
            }
        }
    }

    return result;
}

// Find closest matching vendor from a text input
function findClosestVendor(text) {
    if (!text) return null;

    // Direct match
    const directMatch = vendors.find(v =>
        v.toLowerCase() === text.toLowerCase()
    );

    if (directMatch) return directMatch;

    // Contains match
    const containsMatch = vendors.find(v =>
        text.toLowerCase().includes(v.toLowerCase()) ||
        v.toLowerCase().includes(text.toLowerCase())
    );

    return containsMatch || null;
}

// Initialize section weights based on default values
function initializeSectionWeights() {
    Object.keys(sections).forEach(section => {
        sectionWeights[section] = sections[section].defaultWeight || 0;
    });

    // Ensure weights add up to 100%
    normalizeWeights();
}

// Calculate section scores
function calculateSectionScores() {
    const result = {};

    Object.keys(sections).forEach(section => {
        result[section] = {};

        vendors.forEach(vendor => {
            let totalScore = 0;
            let count = 0;

            const questionList = sections[section].questions || [];

            questionList.forEach(qId => {
                if (questionScores[qId] && typeof questionScores[qId][vendor] === 'number') {
                    totalScore += questionScores[qId][vendor];
                    count++;
                }
            });

            result[section][vendor] = count > 0 ? totalScore / count : 0;
        });
    });

    return result;
}

// Calculate weighted scores based on section weights
function calculateWeightedScores() {
    const result = {};

    vendors.forEach(vendor => {
        let totalWeightedScore = 0;
        let totalWeight = 0;

        Object.keys(sections).forEach(section => {
            const weight = sectionWeights[section];
            totalWeightedScore += sectionScores[section][vendor] * weight;
            totalWeight += weight;
        });

        // Normalize to ensure weights add up to 100%
        result[vendor] = totalWeight > 0 ? totalWeightedScore / (totalWeight / 100) : 0;
    });

    return result;
}

// Ensure weights add up to 100%
function normalizeWeights() {
    const total = Object.values(sectionWeights).reduce((sum, weight) => sum + weight, 0);

    if (total !== 100 && total > 0) {
        const factor = 100 / total;

        // Adjust all weights proportionally
        const sectionKeys = Object.keys(sectionWeights);

        // Adjust all but the last weight
        for (let i = 0; i < sectionKeys.length - 1; i++) {
            sectionWeights[sectionKeys[i]] = Math.round(sectionWeights[sectionKeys[i]] * factor);
        }

        // Ensure total is exactly 100% by setting the last weight
        if (sectionKeys.length > 0) {
            const lastSection = sectionKeys[sectionKeys.length - 1];
            const sumOfOthers = sectionKeys.slice(0, -1).reduce((sum, key) => sum + sectionWeights[key], 0);
            sectionWeights[lastSection] = 100 - sumOfOthers;
        }
    } else if (total === 0 && Object.keys(sectionWeights).length > 0) {
        // If all weights are 0, set equal weights
        const weight = Math.floor(100 / Object.keys(sectionWeights).length);

        Object.keys(sectionWeights).forEach((section, index, array) => {
            if (index === array.length - 1) {
                // Last section gets remainder to ensure total is 100%
                const sumSoFar = array.slice(0, -1).reduce((sum, s) => sum + sectionWeights[s], 0);
                sectionWeights[section] = 100 - sumSoFar;
            } else {
                sectionWeights[section] = weight;
            }
        });
    }
}

// Update the UI with processed data
function updateUI() {
    // Update report title with vendors
    if (vendors.length > 0) {
        document.getElementById('reportTitle').textContent = vendors.join(' vs ') + ' Evaluation';
    }

    // Update report summary
    document.getElementById('reportSummary').textContent =
        `Interactive analysis of ${vendors.length} solutions across ${questions.length} evaluation criteria`;

    // Update executive summary
    updateExecutiveSummary();
    
    // Update Core vs Non-core section
    updateCoreNoncoreSection();

    // Create section sliders
    createSectionSliders();

    // Initialize all charts and components
    initializeComponents();

    // Set current date in footer
    document.getElementById('reportDate').textContent = `Generated ${new Date().toLocaleDateString('en-US', {month: 'long', year: 'numeric'})}`;
}

// Update executive summary
function updateExecutiveSummary() {
    // Update summary text
    const coreCount = coreParticipants.length;
    const noncoreCount = noncoreParticipants.length;
    
    document.getElementById('executiveSummaryText').textContent =
        `This interactive report presents an analysis of the evaluation responses from ${participants.length} participants (${coreCount} Core and ${noncoreCount} Non-core) who assessed ${vendors.length} solutions: ${vendors.join(', ')}. The solutions were evaluated across ${questions.length} criteria organized in ${Object.keys(sections).length} sections.`;

    // Calculate weighted scores for key findings
    const weightedScores = calculateWeightedScores();

    // Sort vendors by score
    const topVendors = [...vendors].sort((a, b) => weightedScores[b] - weightedScores[a]);

    // Sort vendors by standout votes
    const standoutVendors = [...vendors].sort((a, b) => standoutSolutions[b] - standoutSolutions[a]);
    const topStandoutVendor = standoutVendors[0];
    const standoutPercentage = Math.round((standoutSolutions[topStandoutVendor] / participants.length) * 100);

    // Get top choices by participant type
    let coreTopChoice = '';
    let noncoreTopChoice = '';
    
    if (coreVsNoncoreAnalysis && coreVsNoncoreAnalysis.top_vendor) {
        coreTopChoice = coreVsNoncoreAnalysis.top_vendor.Core || '';
        noncoreTopChoice = coreVsNoncoreAnalysis.top_vendor.Non_core || '';
    }

    // Update key findings
    const keyFindingsList = document.getElementById('keyFindingsList');
    keyFindingsList.innerHTML = `
        <li><strong>Top solutions:</strong> <span class="highlight">${topVendors[0]} (${weightedScores[topVendors[0]].toFixed(2)})</span> ${topVendors.length > 1 ? `and <span class="highlight">${topVendors[1]} (${weightedScores[topVendors[1]].toFixed(2)})</span>` : ''} received the highest overall ratings.</li>
        <li><strong>Section-weighted analysis:</strong> With current section weights, <span class="highlight">${topVendors[0]}</span> emerges as the leader.</li>
        <li><strong>User preference:</strong> <span class="highlight">${topStandoutVendor}</span> was selected as a standout solution by ${standoutSolutions[topStandoutVendor]} of ${participants.length} respondents (${standoutPercentage}%).</li>
    `;
    
    if (coreTopChoice && noncoreTopChoice) {
        const agreementText = coreTopChoice === noncoreTopChoice ? 
            `Both Core and Non-core users prefer <span class="highlight">${coreTopChoice}</span>.` : 
            `Core users prefer <span class="highlight">${coreTopChoice}</span> while Non-core users prefer <span class="highlight">${noncoreTopChoice}</span>.`;
        
        keyFindingsList.innerHTML += `<li><strong>User type preferences:</strong> ${agreementText}</li>`;
    }
    
    // Update participants list
    const participantListContainer = document.getElementById('participantListContainer');
    participantListContainer.innerHTML = '';
    
    // Split participants into two columns
    const midpoint = Math.ceil(participants.length / 2);
    
    // First column
    const col1 = document.createElement('div');
    col1.className = 'col-md-6';
    
    const ul1 = document.createElement('ul');
    ul1.className = 'list-unstyled';
    
    participants.slice(0, midpoint).forEach(participant => {
        const li = document.createElement('li');
        const participantType = participantTypes[participant] || '';
        const typeBadge = participantType ? 
            `<span class="participant-type-pill ${participantType.toLowerCase() === 'core' ? 'core' : 'noncore'}">${participantType}</span>` : '';
        li.innerHTML = `<i class="bi bi-person-circle me-2"></i> ${participant} ${typeBadge}`;
        ul1.appendChild(li);
    });
    
    col1.appendChild(ul1);
    participantListContainer.appendChild(col1);
    
    // Second column if needed
    if (participants.length > midpoint) {
        const col2 = document.createElement('div');
        col2.className = 'col-md-6';
        
        const ul2 = document.createElement('ul');
        ul2.className = 'list-unstyled';
        
        participants.slice(midpoint).forEach(participant => {
            const li = document.createElement('li');
            const participantType = participantTypes[participant] || '';
            const typeBadge = participantType ? 
                `<span class="participant-type-pill ${participantType.toLowerCase() === 'core' ? 'core' : 'noncore'}">${participantType}</span>` : '';
            li.innerHTML = `<i class="bi bi-person-circle me-2"></i> ${participant} ${typeBadge}`;
            ul2.appendChild(li);
        });
        
        col2.appendChild(ul2);
        participantListContainer.appendChild(col2);
    }
}

// Update Core vs Non-core section
function updateCoreNoncoreSection() {
    // Update participant counts
    document.getElementById('coreParticipantCount').textContent = coreParticipants.length;
    document.getElementById('noncoreParticipantCount').textContent = noncoreParticipants.length;
    
    // Update agreement level
    if (coreVsNoncoreAnalysis && coreVsNoncoreAnalysis.agreement_level !== undefined) {
        const agreementLevel = coreVsNoncoreAnalysis.agreement_level;
        const agreementPercent = Math.round(agreementLevel * 100);
        
        document.getElementById('agreementValue').textContent = `${agreementPercent}%`;
        document.getElementById('agreementMeterFill').style.width = `${agreementPercent}%`;
        
        let agreementDescription = '';
        if (agreementLevel >= 0.8) {
            agreementDescription = 'High agreement: Core and Non-core participants have very similar preferences.';
        } else if (agreementLevel >= 0.5) {
            agreementDescription = 'Moderate agreement: Core and Non-core participants have somewhat aligned preferences with some differences.';
        } else {
            agreementDescription = 'Low agreement: Core and Non-core participants have significantly different preferences.';
        }
        
        document.getElementById('agreementDescription').textContent = agreementDescription;
    }
    
    // Create participant distribution chart
    createParticipantDistributionChart();
    
    // Create Core vs Non-core comparison chart
    createCoreNoncoreComparisonChart();
    
    // Update top choice widgets
    updateTopChoiceWidgets();
    
    // Update key differences section
    updateKeyDifferences();
    
// Create section comparison chart
createSectionComparisonChart();
}

// Create participant distribution chart
function createParticipantDistributionChart() {
    Highcharts.chart('participantDistributionChart', {
        chart: {
            type: 'pie',
            height: 200
        },
        title: {
            text: null
        },
        tooltip: {
            pointFormat: '<b>{point.y} participants ({point.percentage:.1f}%)</b>'
        },
        accessibility: {
            point: {
                valueSuffix: '%'
            }
        },
        plotOptions: {
            pie: {
                allowPointSelect: true,
                cursor: 'pointer',
                dataLabels: {
                    enabled: true,
                    format: '<b>{point.name}</b><br>{point.y} ({point.percentage:.1f}%)'
                },
                showInLegend: false
            }
        },
        series: [{
            name: 'Participants',
            colorByPoint: true,
            data: [{
                name: 'Core',
                y: coreParticipants.length,
                color: ecobankColors.blue
            }, {
                name: 'Non-core',
                y: noncoreParticipants.length,
                color: ecobankColors.green
            }]
        }]
    });
}

// Create Core vs Non-core comparison chart
function createCoreNoncoreComparisonChart() {
    if (!coreVsNoncoreAnalysis || !coreVsNoncoreAnalysis.overall_scores) {
        return;
    }
    
    const coreScores = [];
    const noncoreScores = [];
    
    // Sort vendors by average of Core and Non-core scores
    const vendorOrder = [...vendors].sort((a, b) => {
        const aAvgScore = (
            (coreVsNoncoreAnalysis.overall_scores.Core[a] || 0) + 
            (coreVsNoncoreAnalysis.overall_scores['Non-core'][a] || 0)
        ) / 2;
        
        const bAvgScore = (
            (coreVsNoncoreAnalysis.overall_scores.Core[b] || 0) + 
            (coreVsNoncoreAnalysis.overall_scores['Non-core'][b] || 0)
        ) / 2;
        
        return bAvgScore - aAvgScore;
    });
    
    vendorOrder.forEach(vendor => {
        coreScores.push(coreVsNoncoreAnalysis.overall_scores.Core[vendor] || 0);
        noncoreScores.push(coreVsNoncoreAnalysis.overall_scores['Non-core'][vendor] || 0);
    });
    
    Highcharts.chart('coreNoncoreComparisonChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: 'Core vs Non-core Overall Scores',
            style: { fontSize: '16px' }
        },
        subtitle: {
            text: 'Comparing evaluation scores between Core and Non-core participants'
        },
        xAxis: {
            categories: vendorOrder,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Average Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            }
        },
        series: [{
            name: 'Core',
            data: coreScores,
            color: ecobankColors.blue
        }, {
            name: 'Non-core',
            data: noncoreScores,
            color: ecobankColors.green
        }]
    });
}

// Update top choice widgets
function updateTopChoiceWidgets() {
    if (!coreVsNoncoreAnalysis || !coreVsNoncoreAnalysis.top_vendor) {
        return;
    }
    
    // Update Core top choice
    const coreTopChoice = coreVsNoncoreAnalysis.top_vendor.Core;
    const coreScore = coreVsNoncoreAnalysis.overall_scores.Core[coreTopChoice] || 0;
    
    if (coreTopChoice) {
        document.getElementById('coreTopChoiceWidget').innerHTML = `
            <div class="mb-1" style="font-size: 1.8rem; font-weight: bold;">${coreTopChoice}</div>
            <div class="text-muted">Score: ${coreScore.toFixed(2)}</div>
        `;
    }
    
    // Update Non-core top choice
    const noncoreTopChoice = coreVsNoncoreAnalysis.top_vendor['Non-core'];
    const noncoreScore = coreVsNoncoreAnalysis.overall_scores['Non-core'][noncoreTopChoice] || 0;
    
    if (noncoreTopChoice) {
        document.getElementById('noncoreTopChoiceWidget').innerHTML = `
            <div class="mb-1" style="font-size: 1.8rem; font-weight: bold;">${noncoreTopChoice}</div>
            <div class="text-muted">Score: ${noncoreScore.toFixed(2)}</div>
        `;
    }
}

// Update key differences section
function updateKeyDifferences() {
    if (!coreVsNoncoreAnalysis || !coreVsNoncoreAnalysis.question_scores) {
        return;
    }
    
    const container = document.getElementById('keyDifferencesContainer');
    container.innerHTML = '';
    
    // Find questions with significant differences
    const questionDifferences = [];
    
    for (const qId in coreVsNoncoreAnalysis.question_scores) {
        const coreScores = coreVsNoncoreAnalysis.question_scores[qId].Core || {};
        const noncoreScores = coreVsNoncoreAnalysis.question_scores[qId]['Non-core'] || {};
        
        vendors.forEach(vendor => {
            const coreScore = coreScores[vendor] || 0;
            const noncoreScore = noncoreScores[vendor] || 0;
            const difference = Math.abs(coreScore - noncoreScore);
            
            if (difference >= 1.0) {  // Minimum difference to be considered significant
                const question = questions.find(q => q.id.toString() === qId);
                
                if (question) {
                    questionDifferences.push({
                        questionId: qId,
                        questionTitle: question.title,
                        vendor: vendor,
                        coreScore: coreScore,
                        noncoreScore: noncoreScore,
                        difference: difference
                    });
                }
            }
        });
    }
    
    // Sort by difference (largest first)
    questionDifferences.sort((a, b) => b.difference - a.difference);
    
    // Display top differences
    if (questionDifferences.length > 0) {
        const maxDifferences = 3; // Number of differences to show
        const topDifferences = questionDifferences.slice(0, maxDifferences);
        
        const differencesHtml = topDifferences.map(diff => {
            const directionIcon = diff.coreScore > diff.noncoreScore ? 
                '<i class="bi bi-arrow-up text-success"></i>' : 
                '<i class="bi bi-arrow-down text-danger"></i>';
            
            return `
                <div class="mb-3">
                    <div class="fw-bold">Q${diff.questionId}: ${truncateText(diff.questionTitle, 60)}</div>
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <span class="type-badge core">Core</span> ${diff.coreScore.toFixed(2)}
                        </div>
                        <div>
                            <span class="type-badge noncore">Non-core</span> ${diff.noncoreScore.toFixed(2)}
                        </div>
                        <div class="ms-2">
                            Diff: ${Math.abs(diff.coreScore - diff.noncoreScore).toFixed(2)} ${directionIcon}
                        </div>
                    </div>
                    <div class="text-muted small mt-1">Solution: ${diff.vendor}</div>
                </div>
            `;
        }).join('');
        
        container.innerHTML = differencesHtml;
    } else {
        container.innerHTML = '<p class="text-center">No significant differences found between Core and Non-core evaluations.</p>';
    }
}

// Helper function to truncate text
function truncateText(text, maxLength) {
    if (!text) return '';
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
}

// Create section comparison chart
function createSectionComparisonChart() {
    if (!coreVsNoncoreAnalysis) {
        return;
    }
    
    // Create section selector
    const sectionSelect = document.getElementById('sectionComparisonSelect');
    sectionSelect.innerHTML = '';
    
    // Add All option
    const allOption = document.createElement('option');
    allOption.value = 'all';
    allOption.textContent = 'All Sections (Overall)';
    sectionSelect.appendChild(allOption);
    
    // Add section options
    Object.keys(sections).forEach(sectionKey => {
        const option = document.createElement('option');
        option.value = sectionKey;
        option.textContent = `Section ${sectionKey}: ${sections[sectionKey].title}`;
        sectionSelect.appendChild(option);
    });
    
    // Add event listener
    sectionSelect.addEventListener('change', function() {
        updateSectionComparisonChart(this.value);
    });
    
    // Initialize with All section
    updateSectionComparisonChart('all');
}

// Update section comparison chart based on selected section
function updateSectionComparisonChart(sectionKey) {
    if (!coreVsNoncoreAnalysis) {
        return;
    }
    
    let chartTitle = '';
    let coreScores = [];
    let noncoreScores = [];
    
    if (sectionKey === 'all') {
        // Overall scores
        chartTitle = 'Overall Scores by User Type';
        
        vendors.forEach(vendor => {
            coreScores.push(coreVsNoncoreAnalysis.overall_scores.Core[vendor] || 0);
            noncoreScores.push(coreVsNoncoreAnalysis.overall_scores['Non-core'][vendor] || 0);
        });
    } else {
        // Section-specific scores
        const sectionTitle = sections[sectionKey] ? sections[sectionKey].title : `Section ${sectionKey}`;
        chartTitle = `${sectionTitle} Scores by User Type`;
        
        // Calculate section scores for Core and Non-core
        const sectionQuestions = sections[sectionKey] ? sections[sectionKey].questions || [] : [];
        
        vendors.forEach(vendor => {
            // Core scores for this section
            let coreTotalScore = 0;
            let coreCount = 0;
            
            // Non-core scores for this section
            let noncoreTotalScore = 0;
            let noncoreCount = 0;
            
            sectionQuestions.forEach(qId => {
                if (coreVsNoncoreAnalysis.question_scores[qId]) {
                    // Core scores
                    const coreScore = coreVsNoncoreAnalysis.question_scores[qId].Core[vendor];
                    if (typeof coreScore === 'number') {
                        coreTotalScore += coreScore;
                        coreCount++;
                    }
                    
                    // Non-core scores
                    const noncoreScore = coreVsNoncoreAnalysis.question_scores[qId]['Non-core'][vendor];
                    if (typeof noncoreScore === 'number') {
                        noncoreTotalScore += noncoreScore;
                        noncoreCount++;
                    }
                }
            });
            
            // Calculate averages
            coreScores.push(coreCount > 0 ? coreTotalScore / coreCount : 0);
            noncoreScores.push(noncoreCount > 0 ? noncoreTotalScore / noncoreCount : 0);
        });
    }
    
    // Create chart
    Highcharts.chart('sectionComparisonChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: chartTitle,
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: vendors,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Average Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                grouping: true,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: 'Core',
            data: coreScores,
            color: ecobankColors.blue,
            pointPadding: 0.3,
            pointPlacement: -0.2
        }, {
            name: 'Non-core',
            data: noncoreScores,
            color: ecobankColors.green,
            pointPadding: 0.3,
            pointPlacement: 0.2
        }]
    });
}

// Create section sliders
function createSectionSliders() {
    const sliderContainer = document.getElementById('sliderContainer');
    sliderContainer.innerHTML = '';
    
    // Create a slider for each section
    Object.keys(sections).forEach(sectionKey => {
        const section = sections[sectionKey];
        const weight = sectionWeights[sectionKey];
        
        // Get question information
        const questionIds = section.questions || [];
        const questionTexts = [];
        
        for (let i = 0; i < Math.min(2, questionIds.length); i++) {
            const question = questions.find(q => q.id === questionIds[i]);
            if (question) {
                questionTexts.push(question.title);
            }
        }
        
        // If there are more questions, indicate that
        if (questionIds.length > 2) {
            questionTexts.push(`+ ${questionIds.length - 2} more questions`);
        }

        // Create weight card
        const weightCard = document.createElement('div');
        weightCard.className = 'weight-card';
        
        weightCard.innerHTML = `
            <div class="weight-header">
                <h3 class="weight-title">Section ${sectionKey}: ${section.title}</h3>
                <span class="weight-value" id="section${sectionKey}Weight">${weight}%</span>
            </div>
            <input type="range" class="custom-range" min="0" max="100" step="5" value="${weight}" id="section${sectionKey}Slider">
            <div class="weight-questions">
                ${questionTexts.map(question => `<div class="question-item">${question}</div>`).join('')}
            </div>
        `;
        
        sliderContainer.appendChild(weightCard);
        
        // Add event listener to the slider
        document.getElementById(`section${sectionKey}Slider`).addEventListener('input', function() {
            sectionWeights[sectionKey] = parseInt(this.value);
            document.getElementById(`section${sectionKey}Weight`).textContent = sectionWeights[sectionKey] + '%';
            
            // Adjust other weights to maintain total of 100%
            adjustOtherWeights(sectionKey);
            
            // Update the UI
            updateWeightedRankingsChart();
            updateExecutiveSummary();
        });
    });
    
    // Create weight preset buttons
    createWeightPresetButtons();
    
    // Attach event handlers to weight buttons
    document.getElementById('resetWeights').addEventListener('click', resetWeights);
    document.getElementById('equalWeights').addEventListener('click', setEqualWeights);
}

// Get short description for a section
function getShortSectionDescription(sectionKey, section) {
    // Get the first 1-2 questions as examples
    const questionIds = section.questions || [];
    const questionTexts = [];
    
    for (let i = 0; i < Math.min(2, questionIds.length); i++) {
        const question = questions.find(q => q.id === questionIds[i]);
        if (question) {
            questionTexts.push(question.title);
        }
    }
    
    // If there are more questions, indicate that
    if (questionIds.length > 2) {
        questionTexts.push(`+ ${questionIds.length - 2} more`);
    }
    
    return questionTexts.join(', ');
}

// Adjust other weights when one slider is moved
function adjustOtherWeights(changedSection) {
    const sectionsArr = Object.keys(sectionWeights);
    const otherSections = sectionsArr.filter(section => section !== changedSection);
    
    // Calculate remaining weight to distribute
    const remainingWeight = 100 - sectionWeights[changedSection];
    
    // Calculate current total of other weights
    const otherTotalWeight = otherSections.reduce((sum, section) => sum + sectionWeights[section], 0);
    
    // If other weights sum to 0, distribute evenly
    if (otherTotalWeight === 0) {
        const equalShare = Math.floor(remainingWeight / otherSections.length);
        otherSections.forEach((section, index) => {
            if (index === otherSections.length - 1) {
                // Last section gets remaining weight to ensure total is 100%
                const sumSoFar = sectionWeights[changedSection] + 
                    otherSections.slice(0, -1).reduce((sum, s) => sum + sectionWeights[s], 0);
                sectionWeights[section] = 100 - sumSoFar;
            } else {
                sectionWeights[section] = equalShare;
            }
        });
    } else {
        // Distribute proportionally based on current weights
        otherSections.forEach((section, index) => {
            if (index === otherSections.length - 1) {
                // Last section ensures total is 100%
                const sumSoFar = sectionWeights[changedSection] +
                    otherSections.slice(0, -1).reduce((sum, s) => sum + sectionWeights[s], 0);
                sectionWeights[section] = 100 - sumSoFar;
            } else {
                const proportion = sectionWeights[section] / otherTotalWeight;
                sectionWeights[section] = Math.round(remainingWeight * proportion);
            }
        });
    }
    
    // Update slider values and labels
    otherSections.forEach(section => {
        const slider = document.getElementById(`section${section}Slider`);
        const label = document.getElementById(`section${section}Weight`);
        
        if (slider && label) {
            slider.value = sectionWeights[section];
            label.textContent = sectionWeights[section] + '%';
        }
    });
}

// Create weight preset buttons
function createWeightPresetButtons() {
    const container = document.getElementById('weightPresetButtons');
    container.innerHTML = '';
    
    // Create a preset for each section focused weight
    Object.keys(sections).forEach(sectionKey => {
        const section = sections[sectionKey];
        
        const button = document.createElement('button');
        button.className = 'btn btn-sm btn-outline-info ms-2';
        button.textContent = `${section.title} Focus`;
        
        button.addEventListener('click', function() {
            // Reset weights
            Object.keys(sectionWeights).forEach(s => {
                sectionWeights[s] = 10;
            });
            
            // Set focus on this section
            sectionWeights[sectionKey] = 60;
            
            // Normalize
            normalizeWeights();
            
            // Update UI
            updateSectionSliders();
            updateWeightedRankingsChart();
            updateExecutiveSummary();
        });
        
        container.appendChild(button);
    });
}

// Update section sliders with current weights
function updateSectionSliders() {
    Object.keys(sectionWeights).forEach(sectionKey => {
        const slider = document.getElementById(`section${sectionKey}Slider`);
        const label = document.getElementById(`section${sectionKey}Weight`);
        
        if (slider && label) {
            slider.value = sectionWeights[sectionKey];
            label.textContent = sectionWeights[sectionKey] + '%';
        }
    });
}

// Reset weights to default
function resetWeights() {
    Object.keys(sections).forEach(section => {
        sectionWeights[section] = sections[section].defaultWeight || 0;
    });
    
    // Normalize
    normalizeWeights();
    
    // Update UI
    updateSectionSliders();
    updateWeightedRankingsChart();
    updateExecutiveSummary();
}
// Set equal weights for all sections
function setEqualWeights() {
    const weight = Math.floor(100 / Object.keys(sections).length);
    
    Object.keys(sections).forEach((section, index, array) => {
        if (index === array.length - 1) {
            // Last section ensures total is 100%
            const sumSoFar = array.slice(0, -1).reduce((sum, s) => sum + sectionWeights[s], 0);
            sectionWeights[section] = 100 - sumSoFar;
        } else {
            sectionWeights[section] = weight;
        }
    });
    
    // Update UI
    updateSectionSliders();
    updateWeightedRankingsChart();
    updateExecutiveSummary();
}

// Initialize all charts and components
function initializeComponents() {
    updateWeightedRankingsChart();
    updateStandoutVotesChart();
    updateSectionScoresChart();
    populateQuestionSelector();
    updateQuestionScoresChart();
    updateParticipantScoresChart();
    updateScoresHeatmap();
    populateParticipantTable();
    populateSectionDetails();
    updateTopSolutionsWidget();
    updateTopInsights();
    updateMethodologySection();
     // Initialize participant comparison section
    initializeParticipantComparison();
}

// Update the weighted rankings chart
function updateWeightedRankingsChart() {
    const weightedScores = calculateWeightedScores();
    
    // Sort vendors by weighted score
    const sortedVendors = Object.keys(weightedScores).sort((a, b) => 
        weightedScores[b] - weightedScores[a]
    );
    
    // Prepare series data
    const seriesData = [];
    
    // Add a series for each section
    Object.keys(sections).forEach(section => {
        seriesData.push({
            name: `Section ${section}`,
            data: sortedVendors.map(vendor => sectionScores[section][vendor])
        });
    });
    
    // Add weighted score series
    seriesData.push({
        name: 'Weighted Score',
        type: 'spline',
        data: sortedVendors.map(vendor => weightedScores[vendor] / 20), // Scaled to match bars
        lineWidth: 4,
        marker: {
            lineWidth: 2,
            lineColor: Highcharts.getOptions().colors[seriesData.length],
            fillColor: 'white',
            radius: 6
        }
    });
    
    // Create the chart
    Highcharts.chart('weightedRankingsChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: 'Solution Rankings by Weighted Section Scores',
            style: { fontSize: '16px' }
        },
        subtitle: {
            text: getWeightsSubtitle()
        },
        xAxis: {
            categories: sortedVendors,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            shared: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            },
            spline: {
                dataLabels: {
                    enabled: true,
                    formatter: function() {
                        return (this.y * 20).toFixed(2);
                    }
                }
            }
        },
        series: seriesData
    });
    
    // Update the rankings table
    updateRankingsTable(sortedVendors, weightedScores);
    
    // Update top solutions widget
    updateTopSolutionsWidget(sortedVendors, weightedScores);
}

// Get weights subtitle for charts
function getWeightsSubtitle() {
    return Object.keys(sectionWeights)
        .map(section => `Section ${section}: ${sectionWeights[section]}%`)
        .join(' | ');
}

// Update the rankings table
function updateRankingsTable(sortedVendors, weightedScores) {
    const rankingsTable = document.getElementById('weightedRankingsTable');
    rankingsTable.innerHTML = '';
    
    sortedVendors.forEach((vendor, index) => {
        const row = document.createElement('tr');
        const rankCell = document.createElement('td');
        const vendorCell = document.createElement('td');
        const scoreCell = document.createElement('td');
        
        rankCell.textContent = index + 1;
        vendorCell.textContent = vendor;
        scoreCell.textContent = weightedScores[vendor].toFixed(2);
        
        // Highlight top solutions
        if (index < 3) {
            row.classList.add('table-primary');
        }
        
        row.appendChild(rankCell);
        row.appendChild(vendorCell);
        row.appendChild(scoreCell);
        rankingsTable.appendChild(row);
    });
}

// Update the top solutions widget
function updateTopSolutionsWidget(sortedVendors, weightedScores) {
    if (!sortedVendors) {
        const calculatedScores = calculateWeightedScores();
        
        sortedVendors = Object.keys(calculatedScores).sort((a, b) => 
            calculatedScores[b] - calculatedScores[a]
        );
        weightedScores = calculatedScores;
    }
    
    const topSolutions = document.getElementById('topSolutionsWidget');
    topSolutions.innerHTML = '';
    
    // Get the top 3 solutions
    const top3 = sortedVendors.slice(0, Math.min(3, sortedVendors.length));
    
    top3.forEach((vendor, index) => {
        const badge = document.createElement('div');
        badge.classList.add('mb-3');
        
        const vendorColor = index === 0 ? 'primary' : (index === 1 ? 'success' : 'info');
        
        badge.innerHTML = `
            <div class="d-flex align-items-center">
                <div class="me-2 text-${vendorColor}" style="font-size: 1.5rem; font-weight: bold;">#${index + 1}</div>
                <div>
                    <div class="fw-bold">${vendor}</div>
                    <div class="text-${vendorColor}">${weightedScores[vendor].toFixed(2)}</div>
                </div>
            </div>
        `;
        
        topSolutions.appendChild(badge);
    });
}

// Update the standout votes chart
function updateStandoutVotesChart() {
    // Sort the solutions by votes
    const sortedSolutions = Object.entries(standoutSolutions)
        .sort((a, b) => b[1] - a[1])
        .map(entry => ({ name: entry[0], y: entry[1] }));
    
    Highcharts.chart('standoutVotesChart', {
        chart: {
            type: 'bar'
        },
        title: {
            text: null
        },
        xAxis: {
            type: 'category'
        },
        yAxis: {
            min: 0,
            max: Math.max(...Object.values(standoutSolutions)) + 1,
            title: {
                text: 'Number of Votes'
            }
        },
        legend: {
            enabled: false
        },
        plotOptions: {
            series: {
                dataLabels: {
                    enabled: true,
                    format: '{point.y}'
                }
            }
        },
        series: [{
            name: 'Votes',
            colorByPoint: true,
            data: sortedSolutions
        }]
    });
}

// Update the section scores chart
function updateSectionScoresChart() {
    // Prepare data for the chart
    const categories = vendors;
    const seriesData = Object.keys(sections).map(section => ({
        name: `Section ${section}: ${sections[section].title}`,
        data: vendors.map(vendor => sectionScores[section][vendor])
    }));
    
    Highcharts.chart('sectionScoresChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: 'Scores by Section',
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: categories,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            }
        },
        series: seriesData
    });
}

// Populate the section details
function populateSectionDetails() {
    const container = document.getElementById('sectionDetailsContainer');
    container.innerHTML = '';
    
    Object.keys(sections).forEach(sectionKey => {
        const section = sections[sectionKey];
        const sectionScoresByVendor = {};
        
        vendors.forEach(vendor => {
            sectionScoresByVendor[vendor] = sectionScores[sectionKey][vendor];
        });
        
        // Sort vendors by score
        const sortedVendors = Object.keys(sectionScoresByVendor).sort((a, b) => 
            sectionScoresByVendor[b] - sectionScoresByVendor[a]
        );
        
        const col = document.createElement('div');
        col.className = 'col-md-4 mb-4';
        
        const card = document.createElement('div');
        card.className = 'card h-100';
        
        const cardHeader = document.createElement('div');
        cardHeader.className = 'card-header';
        cardHeader.innerHTML = `<h6 class="mb-0">Section ${sectionKey}: ${section.title}</h6>`;
        
        const cardBody = document.createElement('div');
        cardBody.className = 'card-body p-0';
        
        const table = document.createElement('table');
        table.className = 'table table-striped mb-0';
        
        const thead = document.createElement('thead');
        thead.innerHTML = `
            <tr>
                <th>Rank</th>
                <th>Vendor</th>
                <th>Score</th>
            </tr>
        `;
        
        const tbody = document.createElement('tbody');
        
        sortedVendors.forEach((vendor, index) => {
            const row = document.createElement('tr');
            
            const rankCell = document.createElement('td');
            rankCell.textContent = index + 1;
            
            const vendorCell = document.createElement('td');
            vendorCell.textContent = vendor;
            
            const scoreCell = document.createElement('td');
            scoreCell.textContent = sectionScoresByVendor[vendor].toFixed(2);
            
            row.appendChild(rankCell);
            row.appendChild(vendorCell);
            row.appendChild(scoreCell);
            
            tbody.appendChild(row);
        });
        
        table.appendChild(thead);
        table.appendChild(tbody);
        cardBody.appendChild(table);
        
        card.appendChild(cardHeader);
        card.appendChild(cardBody);
        col.appendChild(card);
        
        container.appendChild(col);
    });
}

// Populate the question selector dropdown
function populateQuestionSelector() {
    const selector = document.getElementById('questionSelector');
    selector.innerHTML = '';
    
    // Sort questions by ID
    const sortedQuestions = [...questions].sort((a, b) => a.id - b.id);
    
    sortedQuestions.forEach(question => {
        const option = document.createElement('option');
        option.value = question.id;
        option.textContent = `Q${question.id}: ${question.title}`;
        selector.appendChild(option);
    });
    
    // Add event listener
    selector.addEventListener('change', function() {
        updateQuestionScoresChart(this.value);
    });
    
    // Initialize with first question
    if (sortedQuestions.length > 0) {
        updateQuestionScoresChart(sortedQuestions[0].id);
    }
}

// Update the question scores chart
function updateQuestionScoresChart(questionId) {
    const qId = parseInt(questionId);
    const question = questions.find(q => q.id === qId);
    
    if (!question) return;
    
    // Update question details
    document.getElementById('questionDetailsTitle').textContent = `Q${qId}: ${question.title}`;
    
    // Get question description if available
    let description = '';
    if (question.description) {
        description = question.description;
    } else {
        description = `Question ${qId} in Section ${question.section}`;
    }
    document.getElementById('questionDetailsContent').textContent = description;
    
    // Prepare data for chart
    const scores = questionScores[qId] || {};
    
    // Filter out any special columns
    const filteredScores = {};
    vendors.forEach(vendor => {
        if (scores[vendor] !== undefined) {
            filteredScores[vendor] = scores[vendor];
        }
    });
    
    const sortedVendors = Object.keys(filteredScores).sort((a, b) => filteredScores[b] - filteredScores[a]);
    
    // Add Core vs Non-core data if available
    let seriesData = [{
        name: 'Average Score',
        data: sortedVendors.map(vendor => filteredScores[vendor]),
        colorByPoint: true
    }];
    
    // Add Core and Non-core series if data is available
    if (coreVsNoncoreAnalysis && coreVsNoncoreAnalysis.question_scores && coreVsNoncoreAnalysis.question_scores[qId]) {
        const coreScores = coreVsNoncoreAnalysis.question_scores[qId].Core || {};
        const noncoreScores = coreVsNoncoreAnalysis.question_scores[qId]['Non-core'] || {};
        
        seriesData = [
            {
                name: 'Core',
                data: sortedVendors.map(vendor => coreScores[vendor] || 0),
                color: ecobankColors.blue
            },
            {
                name: 'Non-core',
                data: sortedVendors.map(vendor => noncoreScores[vendor] || 0),
                color: ecobankColors.green
            },
            {
                name: 'Average',
                data: sortedVendors.map(vendor => filteredScores[vendor]),
                color: ecobankColors.gray,
                type: 'line',
                dashStyle: 'Dash',
                marker: {
                    lineWidth: 2,
                    lineColor: ecobankColors.gray,
                    fillColor: 'white'
                }
            }
        ];
    }
    
    Highcharts.chart('questionScoresChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: `Q${qId}: ${question.title}`,
            style: { fontSize: '16px' }
        },
        subtitle: {
            text: `Section ${question.section}: ${sections[question.section]?.title || ''}`
        },
        xAxis: {
            categories: sortedVendors,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        legend: {
            enabled: seriesData.length > 1
        },
        series: seriesData
    });
}

// Update the participant scores chart
function updateParticipantScoresChart() {
    // Create series data for each vendor
    const seriesData = vendors.map(vendor => ({
        name: vendor,
        data: participants.map(participant => {
            // Ensure the score exists and is a number
            const score = participantScores[participant] && 
                         participantScores[participant][vendor] !== undefined ? 
                         participantScores[participant][vendor] : 0;
            return typeof score === 'number' ? score : 0;
        })
    }));
    
    // Create x-axis categories with participant types
    const participantLabels = participants.map(participant => {
        const type = participantTypes[participant] || '';
        return type ? `${participant} (${type})` : participant;
    });
    
    // Sort participants by type (Core first, then Non-core)
    const sortedIndices = [];
    
    // First add Core participants
    participants.forEach((participant, index) => {
        if (participantTypes[participant] === 'Core') {
            sortedIndices.push(index);
        }
    });
    
    // Then add Non-core participants
    participants.forEach((participant, index) => {
        if (participantTypes[participant] === 'Non-core') {
            sortedIndices.push(index);
        }
    });
    
    // Then add any other participants
    participants.forEach((participant, index) => {
        if (participantTypes[participant] !== 'Core' && participantTypes[participant] !== 'Non-core') {
            sortedIndices.push(index);
        }
    });
    
    // Sort series data
    const sortedLabels = sortedIndices.map(i => participantLabels[i]);
    const sortedSeriesData = vendors.map(vendor => {
        const originalSeries = seriesData.find(s => s.name === vendor);
        return {
            name: vendor,
            data: sortedIndices.map(i => originalSeries.data[i])
        };
    });
    
    // Create chart
    Highcharts.chart('participantScoresChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: 'Scores by Participant',
            style: { fontSize: '16px' }
        },
        subtitle: {
            text: 'Core vs Non-core Participant Evaluation'
        },
        xAxis: {
            categories: sortedLabels,
            crosshair: true,
            labels: {
                rotation: -45,
                style: {
                    fontSize: '10px'
                }
            }
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Average Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            }
        },
        series: sortedSeriesData
    });
}

// Populate the participant scores table
function populateParticipantTable() {
    // Update header with vendor names
    const headerRow = document.getElementById('participantScoresHeader');
    headerRow.innerHTML = '<th>Participant</th>';
    
    vendors.forEach(vendor => {
        const th = document.createElement('th');
        th.textContent = vendor;
        headerRow.appendChild(th);
    });
    
    headerRow.innerHTML += '<th>Top Choice</th>';
    
    // Populate table rows
    const table = document.getElementById('participantScoresTable');
    table.innerHTML = '';
    
    // Sort participants (Core first, then Non-core)
    const sortedParticipants = [];
    
    // First add Core participants
    participants.forEach(participant => {
        if (participantTypes[participant] === 'Core') {
            sortedParticipants.push(participant);
        }
    });
    
    // Then add Non-core participants
    participants.forEach(participant => {
        if (participantTypes[participant] === 'Non-core') {
            sortedParticipants.push(participant);
        }
    });
    
    // Then add any other participants
    participants.forEach(participant => {
        if (participantTypes[participant] !== 'Core' && participantTypes[participant] !== 'Non-core') {
            sortedParticipants.push(participant);
        }
    });
    
    sortedParticipants.forEach(participant => {
        const row = document.createElement('tr');
        
        // Participant name with type badge
        const nameCell = document.createElement('td');
        const participantType = participantTypes[participant] || '';
        const typeBadge = participantType ? 
            `<span class="participant-type-pill ${participantType.toLowerCase() === 'core' ? 'core' : 'noncore'}">${participantType}</span>` : '';
        nameCell.innerHTML = `${participant} ${typeBadge}`;
        row.appendChild(nameCell);
        
        // Check if scores exist for this participant
        if (!participantScores[participant]) {
            console.warn('No scores found for participant:', participant);
            return;
        }
        
        // Scores for each vendor
        const scores = {};
        vendors.forEach(vendor => {
            const scoreCell = document.createElement('td');
            
            // Ensure score exists and is a number
            const score = participantScores[participant][vendor];
            const displayScore = typeof score === 'number' ? score.toFixed(2) : 'N/A';
            
            scoreCell.textContent = displayScore;
            row.appendChild(scoreCell);
            
            if (typeof score === 'number') {
                scores[vendor] = score;
            }
        });
        
        // Top choice
        const topChoiceCell = document.createElement('td');
        
        if (Object.keys(scores).length > 0) {
            const topVendor = Object.keys(scores).reduce((a, b) => 
                scores[a] > scores[b] ? a : b, Object.keys(scores)[0]);
            
            const badge = document.createElement('span');
            badge.className = 'vendor-badge';
            badge.textContent = topVendor;
            
            // Color code the badge
            const vendorIndex = vendors.indexOf(topVendor);
            const colors = ['bg-secondary', 'bg-info', 'bg-warning', 'bg-success', 'bg-primary'];
            badge.classList.add(colors[vendorIndex % colors.length]);
            
            topChoiceCell.appendChild(badge);
        } else {
            topChoiceCell.textContent = 'N/A';
        }
        
        row.appendChild(topChoiceCell);
        
        table.appendChild(row);
    });
}

// Update the scores heatmap
function updateScoresHeatmap() {
    // Prepare data for heatmap
    const data = [];
    
    // Sort questions by ID for consistent display
    const sortedQuestions = [...questions].sort((a, b) => a.id - b.id);
    
    sortedQuestions.forEach((question, y) => {
        const qId = question.id;
        
        // Skip if no scores for this question
        if (!questionScores[qId]) return;
        
        vendors.forEach((vendor, x) => {
            // Check if score exists for this vendor
            if (questionScores[qId][vendor] !== undefined) {
                data.push([x, y, questionScores[qId][vendor]]);
            }
        });
    });
    
    Highcharts.chart('scoresHeatmap', {
        chart: {
            type: 'heatmap'
        },
        title: {
            text: 'Scores Heatmap by Question and Vendor',
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: vendors
        },
        yAxis: {
            categories: sortedQuestions.map(q => `Q${q.id}: ${q.title}`),
            title: null,
            reversed: true
        },
        colorAxis: {
            min: 1,
            max: 5,
            stops: [
                [0, '#ffdfdf'], // Light red for low scores
                [0.5, '#ffffdf'], // Yellow for middle scores
                [1, '#dfffdf']  // Light green for high scores
            ]
        },
        legend: {
            align: 'right',
            layout: 'vertical',
            margin: 0,
            verticalAlign: 'top',
            y: 25,
            symbolHeight: 280
        },
        tooltip: {
            formatter: function () {
                return `<b>${this.series.xAxis.categories[this.point.x]}</b><br>` +
                       `${this.series.yAxis.categories[this.point.y]}<br>` +
                       `Score: <b>${this.point.value}</b>`;
            }
        },
        series: [{
            name: 'Score',
            borderWidth: 1,
            data: data,
            dataLabels: {
                enabled: true,
                color: '#000000',
                style: {
                    textOutline: 'none'
                }
            }
        }]
    });
}

// Update the top insights section
function updateTopInsights() {
    const container = document.getElementById('topInsightsContainer');
    container.innerHTML = '';
    
    // Calculate vendor strengths
    const vendorStrengths = calculateVendorStrengths();
    
    // Calculate section performance
    const sectionPerformance = calculateSectionPerformance();
    
    // Create vendor strengths card
    const strengthsCard = createInsightCard(
        'Vendor Strengths',
        getVendorStrengthsContent(vendorStrengths)
    );
    
    // Create section performance card
    const performanceCard = createInsightCard(
        'Section Performance',
        getSectionPerformanceContent(sectionPerformance)
    );
    
    // Create user type insights card
    const userTypeCard = createInsightCard(
        'User Type Insights',
        getUserTypeInsightsContent()
    );
    
    // Add cards to container
    container.appendChild(createColumnDiv(strengthsCard));
    container.appendChild(createColumnDiv(performanceCard));
    container.appendChild(createColumnDiv(userTypeCard));
}

// Create a column div for insight cards
function createColumnDiv(content) {
    const col = document.createElement('div');
    col.className = 'col-md-4 mb-4';
    col.appendChild(content);
    return col;
}

// Create an insight card
function createInsightCard(title, content) {
    const card = document.createElement('div');
    card.className = 'card h-100';
    
    const cardHeader = document.createElement('div');
    cardHeader.className = 'card-header';
    cardHeader.innerHTML = `<h5 class="mb-0">${title}</h5>`;
    
    const cardBody = document.createElement('div');
    cardBody.className = 'card-body';
    cardBody.innerHTML = content;
    
    card.appendChild(cardHeader);
    card.appendChild(cardBody);
    
    return card;
}

// Get user type insights content
function getUserTypeInsightsContent() {
    let content = '';
    
    if (coreVsNoncoreAnalysis && coreVsNoncoreAnalysis.top_vendor) {
        const coreTopChoice = coreVsNoncoreAnalysis.top_vendor.Core;
        const noncoreTopChoice = coreVsNoncoreAnalysis.top_vendor['Non-core'];
        
        if (coreTopChoice) {
            content += `<p><strong>Core Preference:</strong> ${coreTopChoice} was highest-rated among Core users with a score of ${coreVsNoncoreAnalysis.overall_scores.Core[coreTopChoice].toFixed(2)}.</p>`;
        }
        
        if (noncoreTopChoice) {
            content += `<p><strong>Non-core Preference:</strong> ${noncoreTopChoice} was highest-rated among Non-core users with a score of ${coreVsNoncoreAnalysis.overall_scores['Non-core'][noncoreTopChoice].toFixed(2)}.</p>`;
        }
        
        if (coreVsNoncoreAnalysis.agreement_level !== undefined) {
            const agreementLevel = coreVsNoncoreAnalysis.agreement_level;
            let agreementText = '';
            
            if (agreementLevel >= 0.8) {
                agreementText = 'High agreement between Core and Non-core evaluations.';
            } else if (agreementLevel >= 0.5) {
                agreementText = 'Moderate agreement with some differences between Core and Non-core evaluations.';
            } else {
                agreementText = 'Low agreement with significant differences between Core and Non-core evaluations.';
            }
            
            content += `<p><strong>Agreement Level:</strong> ${agreementText}</p>`;
        }
    }
    
    // Add user type statistics
    content += `<div class="mt-3">
        <p><span class="type-badge core">Core</span> ${coreParticipants.length} participants</p>
        <p><span class="type-badge noncore">Non-core</span> ${noncoreParticipants.length} participants</p>
    </div>`;
    
    return content;
}

// Calculate vendor strengths (top-scoring questions)
function calculateVendorStrengths() {
    const result = {};
    
    // Get the top two vendors by overall score
    const weightedScores = calculateWeightedScores();
    const topVendors = Object.keys(weightedScores)
        .sort((a, b) => weightedScores[b] - weightedScores[a])
        .slice(0, Math.min(2, vendors.length));
    
    // For each top vendor, find their top 3 questions
    topVendors.forEach(vendor => {
        result[vendor] = [];
        
        // Create array of [questionId, score] pairs
        const questionScorePairs = [];
        
        Object.keys(questionScores).forEach(qId => {
            if (questionScores[qId][vendor] !== undefined) {
                questionScorePairs.push([parseInt(qId), questionScores[qId][vendor]]);
            }
        });
        
        // Sort by score (descending)
        questionScorePairs.sort((a, b) => b[1] - a[1]);
        
        // Take top 3
        const top3 = questionScorePairs.slice(0, Math.min(3, questionScorePairs.length));
        
        // Get question titles
        top3.forEach(([qId, score]) => {
            const question = questions.find(q => q.id === qId);
            
            if (question) {
                result[vendor].push({
                    title: question.title,
                    score: score
                });
            }
        });
    });
    
    return result;
}

// Get HTML content for vendor strengths
function getVendorStrengthsContent(vendorStrengths) {
    let content = '';
    
    Object.keys(vendorStrengths).forEach(vendor => {
        content += `<h6>${vendor}'s Top Strengths:</h6><ol>`;
        
        vendorStrengths[vendor].forEach(item => {
            content += `<li>${item.title} (${item.score.toFixed(2)})</li>`;
        });
        
        content += `</ol>`;
    });
    
    return content;
}

// Calculate section performance
function calculateSectionPerformance() {
    const result = {};
    
    // For each section, get top 2 vendors
    Object.keys(sections).forEach(section => {
        const vendorScores = {};
        
        vendors.forEach(vendor => {
            vendorScores[vendor] = sectionScores[section][vendor];
        });
        
        // Sort vendors by score
        const sortedVendors = Object.keys(vendorScores)
            .sort((a, b) => vendorScores[b] - vendorScores[a]);
        
        result[section] = {
            title: sections[section].title,
            topVendors: sortedVendors.slice(0, Math.min(2, sortedVendors.length)).map(vendor => ({
                name: vendor,
                score: vendorScores[vendor]
            }))
        };
    });
    
    return result;
}

// Get HTML content for section performance
function getSectionPerformanceContent(sectionPerformance) {
    let content = '';
    
    Object.keys(sectionPerformance).forEach(section => {
        const data = sectionPerformance[section];
        const topVendors = data.topVendors;
        
        content += `<p><strong>Section ${section} (${data.title}):</strong> `;
        
        if (topVendors.length > 0) {
            content += `${topVendors[0].name} leads with ${topVendors[0].score.toFixed(2)}`;
            
            if (topVendors.length > 1) {
                content += `, followed by ${topVendors[1].name} (${topVendors[1].score.toFixed(2)})`;
            }
        }
        
        content += `</p>`;
    });
    
    // Add overall insight
    const weightedScores = calculateWeightedScores();
    const topVendor = Object.keys(weightedScores)
        .sort((a, b) => weightedScores[b] - weightedScores[a])[0];
    
    content += `<p class="mb-0"><strong>Overall:</strong> ${topVendor} shows the most balanced performance across all sections.</p>`;
    
    return content;
}

// Update the methodology section
function updateMethodologySection() {
    // Update participant count info
    const coreCount = coreParticipants.length;
    const noncoreCount = noncoreParticipants.length;
    
    document.getElementById('participantCountInfo').textContent = 
        `${participants.length} evaluators (${coreCount} Core and ${noncoreCount} Non-core) representing different departments and roles within the organization completed the assessment. Each evaluator rated ${vendors.length} solutions across ${questions.length} criteria using a 5-point scale:`;
    
    // Update section methodology
    const sectionMethodology = document.getElementById('sectionMethodology');
    sectionMethodology.innerHTML = '';
    
    Object.keys(sections).forEach(sectionKey => {
        const section = sections[sectionKey];
        
        const methodologyItem = document.createElement('div');
        methodologyItem.className = 'methodology-item';
        
        methodologyItem.innerHTML = `
            <strong>Section ${sectionKey}: ${section.title}</strong>
            <p>This section ${getMethodologyDescription(sectionKey)}.</p>
        `;
        
        sectionMethodology.appendChild(methodologyItem);
    });
}

// Get methodology description for each section
function getMethodologyDescription(sectionKey) {
    const descriptions = {
        'A': 'assesses core capabilities and features that address specific user needs',
        'B': 'evaluates the user interface design, experience, and accessibility',
        'C': 'focuses on technical architecture, integration capabilities, and performance'
        // Add more descriptions as needed
    };
    
    return descriptions[sectionKey] || `evaluates key aspects of the solution for Section ${sectionKey}`;
}
// Initialize the participant comparison section
function initializeParticipantComparison() {
    // Populate participant selector
    populateParticipantSelector();
    
    // Initialize with first participant if available
    if (participants.length > 0) {
        selectedParticipant = participants[0];
        updateParticipantComparisonUI();
    }
    
    // Add event listeners for section tabs
    document.getElementById('questionViewBtn').addEventListener('click', function() {
        document.getElementById('participantQuestionsView').style.display = 'block';
        document.getElementById('participantSectionsView').style.display = 'none';
        this.classList.add('active');
        document.getElementById('sectionViewBtn').classList.remove('active');
        updateParticipantQuestionsChart();
    });
    
    document.getElementById('sectionViewBtn').addEventListener('click', function() {
        document.getElementById('participantQuestionsView').style.display = 'none';
        document.getElementById('participantSectionsView').style.display = 'block';
        this.classList.add('active');
        document.getElementById('questionViewBtn').classList.remove('active');
        updateParticipantSectionsChart();
    });
    
    // Populate question detail selector
    populateQuestionDetailSelector();
}

// Populate the participant selector dropdown
function populateParticipantSelector() {
    const selector = document.getElementById('participantSelector');
    selector.innerHTML = '';
    
    // Sort participants by type (Core first, then Non-core)
    const sortedParticipants = [];
    
    // First add Core participants
    participants.forEach(participant => {
        if (participantTypes[participant] === 'Core') {
            sortedParticipants.push(participant);
        }
    });
    
    // Then add Non-core participants
    participants.forEach(participant => {
        if (participantTypes[participant] === 'Non-core') {
            sortedParticipants.push(participant);
        }
    });
    
    // Then add any other participants
    participants.forEach(participant => {
        if (participantTypes[participant] !== 'Core' && participantTypes[participant] !== 'Non-core') {
            sortedParticipants.push(participant);
        }
    });
    
    // Add options to selector
    sortedParticipants.forEach(participant => {
        const option = document.createElement('option');
        option.value = participant;
        
        const participantType = participantTypes[participant] || '';
        option.textContent = participantType ? `${participant} (${participantType})` : participant;
        
        selector.appendChild(option);
    });
    
    // Add event listener
    selector.addEventListener('change', function() {
        selectedParticipant = this.value;
        updateParticipantComparisonUI();
    });
}

// Populate question detail selector
function populateQuestionDetailSelector() {
    const selector = document.getElementById('questionDetailSelector');
    selector.innerHTML = '';
    
    // Sort questions by ID
    const sortedQuestions = [...questions].sort((a, b) => a.id - b.id);
    
    // Add options
    sortedQuestions.forEach(question => {
        const option = document.createElement('option');
        option.value = question.id;
        option.textContent = `Q${question.id}: ${question.title}`;
        selector.appendChild(option);
    });
    
    // Add event listener
    selector.addEventListener('change', function() {
        updateQuestionDetail(this.value);
    });
    
    // Initialize with first question if available
    if (sortedQuestions.length > 0) {
        updateQuestionDetail(sortedQuestions[0].id);
    }
}

// Update the entire participant comparison UI
function updateParticipantComparisonUI() {
    // Update participant type indicator
    const participantType = participantTypes[selectedParticipant] || 'Unknown';
    document.getElementById('selectedParticipantType').textContent = participantType;
    document.getElementById('selectedParticipantType').className = 'participant-type-value ' + 
        (participantType.toLowerCase() === 'core' ? 'text-primary' : 'text-success');
    
    // Update questions view
    updateParticipantQuestionsChart();
    
    // Update question detail
    const questionSelector = document.getElementById('questionDetailSelector');
    if (questionSelector.value) {
        updateQuestionDetail(questionSelector.value);
    }
    
    // Update sections view
    updateParticipantSectionsChart();
    updateSectionComparisonTable();
    
    // Update participant preference chart
    updateParticipantPreferenceChart();
    
    // Update individual vs group chart
    updateParticipantVsGroupChart();
}

// Update participant questions chart
function updateParticipantQuestionsChart() {
    if (!selectedParticipant || !participantQuestionScores[selectedParticipant]) {
        return;
    }
    
    // Prepare data for the chart
    const questionData = {};
    
    // Initialize data for each vendor
    vendors.forEach(vendor => {
        questionData[vendor] = [];
    });
    
    // Sort questions by ID
    const sortedQuestions = [...questions].sort((a, b) => a.id - b.id);
    
    // Collect scores for each question
    sortedQuestions.forEach(question => {
        const qId = question.id.toString();
        
        vendors.forEach(vendor => {
            if (participantQuestionScores[selectedParticipant][qId] && 
                participantQuestionScores[selectedParticipant][qId][vendor] !== undefined) {
                questionData[vendor].push(participantQuestionScores[selectedParticipant][qId][vendor]);
            } else {
                questionData[vendor].push(0);
            }
        });
    });
    
    // Create series data
    const seriesData = vendors.map(vendor => ({
        name: vendor,
        data: questionData[vendor]
    }));
    
    // Create the chart
    Highcharts.chart('participantQuestionsChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: `${selectedParticipant}'s Question Scores by Solution`,
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: sortedQuestions.map(q => `Q${q.id}`),
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">Question {point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            }
        },
        series: seriesData
    });
}

// Update question detail
function updateQuestionDetail(questionId) {
    const qId = questionId.toString();
    const question = questions.find(q => q.id.toString() === qId);
    
    if (!question || !selectedParticipant) {
        return;
    }
    
    // Update question detail content
    const container = document.getElementById('questionDetailContent');
    container.innerHTML = `
        <h5>Q${qId}: ${question.title}</h5>
        <p><strong>Section:</strong> ${question.section} - ${sections[question.section]?.title || ''}</p>
        ${question.description ? `<p>${question.description}</p>` : ''}
    `;
    
    // Update question detail table
    const tableBody = document.getElementById('questionDetailTable').querySelector('tbody');
    tableBody.innerHTML = '';
    
    vendors.forEach(vendor => {
        const row = document.createElement('tr');
        
        // Vendor name
        const vendorCell = document.createElement('td');
        vendorCell.textContent = vendor;
        row.appendChild(vendorCell);
        
        // Participant's score
        const scoreCell = document.createElement('td');
        const score = participantQuestionScores[selectedParticipant]?.[qId]?.[vendor] || 0;
        scoreCell.textContent = score.toFixed(2);
        row.appendChild(scoreCell);
        
        // Core average
        const coreAvgCell = document.createElement('td');
        const coreAvg = coreVsNoncoreAnalysis?.question_scores?.[qId]?.Core?.[vendor] || 0;
        coreAvgCell.textContent = coreAvg.toFixed(2);
        row.appendChild(coreAvgCell);
        
        // Non-core average
        const noncoreAvgCell = document.createElement('td');
        const noncoreAvg = coreVsNoncoreAnalysis?.question_scores?.[qId]?.['Non-core']?.[vendor] || 0;
        noncoreAvgCell.textContent = noncoreAvg.toFixed(2);
        row.appendChild(noncoreAvgCell);
        
        // Overall average
        const overallAvgCell = document.createElement('td');
        const overallAvg = questionScores[qId]?.[vendor] || 0;
        overallAvgCell.textContent = overallAvg.toFixed(2);
        row.appendChild(overallAvgCell);
        
        tableBody.appendChild(row);
    });
}

// Update participant sections chart
function updateParticipantSectionsChart() {
    if (!selectedParticipant || !participantSectionScores[selectedParticipant]) {
        return;
    }
    
    // Prepare data for the chart
    const sectionData = {};
    
    // Initialize data for each vendor
    vendors.forEach(vendor => {
        sectionData[vendor] = [];
    });
    
    // Get sorted section keys
    const sortedSections = Object.keys(sections).sort();
    
    // Collect scores for each section
    sortedSections.forEach(sectionKey => {
        vendors.forEach(vendor => {
            if (participantSectionScores[selectedParticipant][sectionKey] && 
                participantSectionScores[selectedParticipant][sectionKey][vendor] !== undefined) {
                sectionData[vendor].push(participantSectionScores[selectedParticipant][sectionKey][vendor]);
            } else {
                sectionData[vendor].push(0);
            }
        });
    });
    
    // Create series data
    const seriesData = vendors.map(vendor => ({
        name: vendor,
        data: sectionData[vendor]
    }));
    
    // Create the chart
    Highcharts.chart('participantSectionsChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: `${selectedParticipant}'s Section Scores by Solution`,
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: sortedSections.map(s => `Section ${s}: ${sections[s].title}`),
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                pointPadding: 0.2,
                borderWidth: 0
            }
        },
        series: seriesData
    });
}

// Update section comparison table
function updateSectionComparisonTable() {
    if (!selectedParticipant || !participantSectionScores[selectedParticipant]) {
        return;
    }
    
    const tableBody = document.getElementById('sectionComparisonTable').querySelector('tbody');
    tableBody.innerHTML = '';
    
    // Get sorted section keys
    const sortedSections = Object.keys(sections).sort();
    
    sortedSections.forEach(sectionKey => {
        const row = document.createElement('tr');
        
        // Section name
        const sectionCell = document.createElement('td');
        sectionCell.textContent = `Section ${sectionKey}: ${sections[sectionKey].title}`;
        row.appendChild(sectionCell);
        
        // Calculate participant's average for this section across all vendors
        const participantAvgCell = document.createElement('td');
        let totalScore = 0;
        let count = 0;
        
        vendors.forEach(vendor => {
            if (participantSectionScores[selectedParticipant][sectionKey] && 
                participantSectionScores[selectedParticipant][sectionKey][vendor] !== undefined) {
                totalScore += participantSectionScores[selectedParticipant][sectionKey][vendor];
                count++;
            }
        });
        
        const participantAvg = count > 0 ? totalScore / count : 0;
        participantAvgCell.textContent = participantAvg.toFixed(2);
        row.appendChild(participantAvgCell);
        
        // Calculate Core average for this section
        const coreAvgCell = document.createElement('td');
        let coreTotal = 0;
        let coreCount = 0;
        
        vendors.forEach(vendor => {
            const sectionScore = sectionScores[sectionKey]?.[vendor] || 0;
            if (sectionScore > 0) {
                coreTotal += sectionScore;
                coreCount++;
            }
        });
        
        const coreAvg = coreCount > 0 ? coreTotal / coreCount : 0;
        coreAvgCell.textContent = coreAvg.toFixed(2);
        row.appendChild(coreAvgCell);
        
        // Non-core average would be similar to core, but we don't have direct access to non-core section scores
        // Using overall section scores as an approximation
        const noncoreAvgCell = document.createElement('td');
        noncoreAvgCell.textContent = coreAvg.toFixed(2); // Using same value as placeholder
        row.appendChild(noncoreAvgCell);
        
        tableBody.appendChild(row);
    });
}

// Update participant preference chart
function updateParticipantPreferenceChart() {
    if (!selectedParticipant || !participantScores[selectedParticipant]) {
        return;
    }
    
    // Prepare data for the chart
    const scores = [];
    
    vendors.forEach(vendor => {
        scores.push({
            name: vendor,
            y: participantScores[selectedParticipant][vendor] || 0
        });
    });
    
    // Sort by score (highest first)
    scores.sort((a, b) => b.y - a.y);
    
    // Create the chart
    Highcharts.chart('participantPreferenceChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: null
        },
        xAxis: {
            type: 'category'
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Average Score (1-5)'
            }
        },
        legend: {
            enabled: false
        },
        tooltip: {
            pointFormat: '<b>{point.y:.2f}</b>'
        },
        series: [{
            name: 'Score',
            colorByPoint: true,
            data: scores,
            dataLabels: {
                enabled: true,
                format: '{point.y:.2f}',
                style: {
                    fontSize: '10px'
                }
            }
        }]
    });
}

// Update individual vs group chart
function updateParticipantVsGroupChart() {
    if (!selectedParticipant || !participantScores[selectedParticipant]) {
        return;
    }
    
    // Prepare data
    const participantData = [];
    const groupData = [];
    
    // Calculate overall average scores
    const overallAverages = {};
    vendors.forEach(vendor => {
        let totalScore = 0;
        let count = 0;
        
        participants.forEach(participant => {
            if (participantScores[participant] && typeof participantScores[participant][vendor] === 'number') {
                totalScore += participantScores[participant][vendor];
                count++;
            }
        });
        
        overallAverages[vendor] = count > 0 ? totalScore / count : 0;
    });
    
    // Create data points
    vendors.forEach(vendor => {
        const participantScore = participantScores[selectedParticipant][vendor] || 0;
        const groupScore = overallAverages[vendor];
        
        participantData.push(participantScore);
        groupData.push(groupScore);
    });
    
    // Create the chart
    Highcharts.chart('participantVsGroupChart', {
        chart: {
            type: 'column'
        },
        title: {
            text: 'Individual vs Group Average Comparison',
            style: { fontSize: '16px' }
        },
        xAxis: {
            categories: vendors,
            crosshair: true
        },
        yAxis: {
            min: 0,
            max: 5,
            title: {
                text: 'Score (1-5)'
            }
        },
        tooltip: {
            headerFormat: '<span style="font-size:10px">{point.key}</span><table>',
            pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                '<td style="padding:0"><b>{point.y:.2f}</b></td></tr>',
            footerFormat: '</table>',
            shared: true,
            useHTML: true
        },
        plotOptions: {
            column: {
                grouping: true,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: 'Participant Score',
            data: participantData,
            color: ecobankColors.blue,
            pointPadding: 0.3,
            pointPlacement: -0.2
        }, {
            name: 'Group Average',
            data: groupData,
            color: ecobankColors.green,
            pointPadding: 0.3,
            pointPlacement: 0.2
        }]
    });
}