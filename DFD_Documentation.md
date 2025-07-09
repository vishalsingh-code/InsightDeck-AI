# Data Flow Diagram (DFD) - CSV-to-PowerPoint AI Analyzer

## Project Overview
The CSV-to-PowerPoint AI Analyzer is a system that transforms CSV data files into comprehensive PowerPoint presentations with AI-powered insights and professional visualizations.

## DFD Level 0 - Context Diagram

```
                    ┌─────────────────────────────────────────────────────────┐
                    │                                                         │
                    │        CSV-to-PowerPoint                                │
           CSV ────▶│         AI Analyzer                                     │────▶ PowerPoint
          File      │                                                         │     Presentation
                    │                                                         │
     OpenAI API ────▶│                                                         │
      Response      │                                                         │
                    │                                                         │
    User Config ────▶│                                                         │
                    │                                                         │
                    └─────────────────────────────────────────────────────────┘
                                        │
                                        ▼
                                Enhanced Slide Content
                                (Temporary Files)
```

### External Entities:
- **User**: Provides CSV file and configures cleaning options, receives PowerPoint presentation
- **OpenAI API**: Provides AI-generated insights and analysis
- **File System**: Stores temporary chart images, enhanced temp files, and output files

---

## DFD Level 1 - Detailed Process Flow

```
                            CSV File
                               │
                               ▼
                    ┌─────────────────────┐
                    │    1.0 LOAD &       │
                    │   CLEAN DATA        │◄─── Data Quality Rules
                    │                     │
                    └──────────┬──────────┘
                               │
                        Cleaned Data
                               │
                               ▼
                    ┌─────────────────────┐
                    │   2.0 ANALYZE       │
                    │   CSV DATA          │◄─── Statistical Methods
                    │                     │
                    └──────────┬──────────┘
                               │
                        Data Analysis
                               │
                               ▼
                    ┌─────────────────────┐      API Request
                    │   3.0 GENERATE      │─────────────────┐
                    │  AI INSIGHTS        │                 │
                    │                     │◄─────────────────┘
                    └──────────┬──────────┘      AI Response
                               │
                        Enhanced Insights
                               │
                               ▼
                    ┌─────────────────────┐
                    │   4.0 CREATE        │
                    │   CHARTS            │
                    │                     │
                    └──────────┬──────────┘
                               │
                        Chart Images
                               │
                               ▼
                    ┌─────────────────────┐      Chart Files
                    │   5.0 BUILD         │◄─────────────────
                    │  PRESENTATION       │
                    │                     │
                    └──────────┬──────────┘
                               │
                               ▼
                      PowerPoint File
```

---

## DFD Level 2 - Detailed Sub-Processes

### 2.1 Process 1.0 - LOAD & ANALYZE CSV

```
    CSV File
       │
       ▼
┌─────────────────┐     Raw Data      ┌─────────────────┐
│   1.1 LOAD      │─────────────────▶│  D1: RAW_DATA   │
│   CSV DATA      │                   │                 │
└─────────────────┘                   └─────────────────┘
       │                                      │
       │                                      ▼
       │                              ┌─────────────────┐
       │                              │   1.2 BASIC     │
       │                              │   ANALYSIS      │
       │                              │                 │
       │                              └─────────┬───────┘
       │                                        │
       │                                Basic Stats
       │                                        │
       │                                        ▼
       │                              ┌─────────────────┐
       │                              │   1.3 ADVANCED  │
       │                              │   ANALYSIS      │
       │                              │                 │
       │                              └─────────┬───────┘
       │                                        │
       │                              Comprehensive Analysis
       │                                        │
       │                                        ▼
       │                              ┌─────────────────┐
       └─────────────────────────────▶│ D2: ANALYSIS    │
                                      │    RESULTS      │
                                      └─────────────────┘
```

### 2.2 Process 2.0 - GENERATE AI INSIGHTS

```
Analysis Results
       │
       ▼
┌─────────────────┐   Data Summary    ┌─────────────────┐
│   2.1 BUILD     │─────────────────▶│   2.2 CALL      │
│  PROMPT         │                   │  OPENAI API     │
└─────────────────┘                   └─────────┬───────┘
       │                                        │
       │                                API Response
       │                                        │
       │                                        ▼
       │                              ┌─────────────────┐
       │                              │   2.3 PARSE     │
       │                              │   RESPONSE      │
       │                              │                 │
       │                              └─────────┬───────┘
       │                                        │
       │                                Parsed Insights
       │                                        │
       │                                        ▼
       │                              ┌─────────────────┐
       │                              │   2.4 VALIDATE  │
       │                              │  & ENHANCE      │
       │                              │                 │
       │                              └─────────┬───────┘
       │                                        │
       │                              Enhanced Insights
       │                                        │
       │                                        ▼
       └─────────────────────────────▶│ D3: STRUCTURED  │
                                      │    INSIGHTS     │
                                      └─────────────────┘
```

### 2.3 Process 3.0 - CREATE CHARTS

```
Structured Insights + Analysis Results
                │
                ▼
        ┌─────────────────┐
        │   3.1 SELECT    │
        │  CHART TYPES    │
        │                 │
        └─────────┬───────┘
                  │
           Chart Specifications
                  │
                  ▼
        ┌─────────────────┐     Chart Data     ┌─────────────────┐
        │   3.2 PREPARE   │──────────────────▶│   3.3 RENDER    │
        │   DATA FOR      │                   │   CHARTS        │
        │   CHARTS        │                   │                 │
        └─────────────────┘                   └─────────┬───────┘
                                                        │
                                                Chart Images
                                                        │
                                                        ▼
                                              ┌─────────────────┐
                                              │  D4: CHART      │
                                              │     FILES       │
                                              └─────────────────┘
```

### 2.4 Process 4.0 - BUILD PRESENTATION

```
Structured Insights + Chart Files
                │
                ▼
        ┌─────────────────┐
        │   4.1 CREATE    │
        │  PRESENTATION   │
        │   STRUCTURE     │
        └─────────┬───────┘
                  │
            Slide Structure
                  │
                  ▼
        ┌─────────────────┐      Slide Data     ┌─────────────────┐
        │   4.2 CREATE    │────────────────────▶│   4.3 ADD       │
        │   SLIDES        │                     │   CHARTS        │
        │                 │                     │                 │
        └─────────────────┘                     └─────────┬───────┘
                                                          │
                                                 Complete Slides
                                                          │
                                                          ▼
                                                ┌─────────────────┐
                                                │   4.4 SAVE      │
                                                │  PRESENTATION   │
                                                │                 │
                                                └─────────┬───────┘
                                                          │
                                                          ▼
                                                PowerPoint File
```

---

## Data Stores

### D1: RAW_DATA
- **Content**: Original CSV data loaded into pandas DataFrame
- **Structure**: Rows and columns as read from CSV
- **Access**: Read-only after initial load

### D2: CLEANED_DATA
- **Content**: Cleaned and validated data after quality processing:
  - Multi-encoding detection and correction
  - Missing value handling and outlier removal
  - Data type optimization and normalization
  - Business rule validation
- **Structure**: Cleaned pandas DataFrame with quality metrics
- **Access**: Read for analysis and processing

### D3: ANALYSIS_RESULTS
- **Content**: Comprehensive data analysis including:
  - Statistical summaries (mean, median, std, etc.)
  - Data quality metrics (missing values, duplicates)
  - Correlation matrices and outlier detection
  - Pattern identification and cleaning summary
  - Column classifications and data completeness
- **Structure**: Enhanced nested dictionary with analysis categories
- **Access**: Read for AI insight generation and chart creation

### D4: STRUCTURED_INSIGHTS
- **Content**: AI-generated and enhanced insights including:
  - Presentation title and structure
  - Key findings and insights (5-8 bullet points per slide)
  - Chart recommendations with proper y-columns
  - Slide content specifications with detailed bullet points
- **Structure**: JSON-like dictionary with enhanced presentation metadata
- **Access**: Read for presentation building

### D5: CHART_FILES
- **Content**: Temporary PNG files for charts
- **Structure**: File paths to chart images
- **Access**: Read for presentation building, deleted after completion
- **Types**: Bar charts, pie charts, line charts, scatter plots, heatmaps

---

## Data Flows Description

### Primary Data Flows:

1. **CSV File → Raw Data**: Initial file reading and DataFrame creation
2. **Raw Data → Data Analysis**: Statistical and structural analysis
3. **Data Analysis → AI Prompt**: Formatted summary for AI processing
4. **AI Response → Structured Insights**: Parsed and validated AI output
5. **Chart Specifications → Chart Images**: Matplotlib/Seaborn visualization generation
6. **Structured Content → PowerPoint Slides**: python-pptx slide creation
7. **Complete Presentation → Output File**: Final PPTX file generation

### Control Flows:

1. **API Key Validation**: Ensures OpenAI access before processing
2. **Data Quality Checks**: Validates CSV structure and content
3. **Fallback Mechanisms**: Handles AI failures with smart defaults
4. **Error Handling**: Manages exceptions throughout the pipeline

### Temporary Flows:

1. **Chart File Management**: Creation and cleanup of temporary images
2. **Memory Management**: DataFrame and analysis result handling

---

## Process Specifications

### 1.0 LOAD & CLEAN DATA
- **Input**: CSV file path
- **Output**: Cleaned and validated data
- **Processing**: 
  - Multi-encoding detection (UTF-8, UTF-8-sig, Latin1, ISO-8859-1)
  - BOM handling and column name cleaning
  - Missing value detection and intelligent handling
  - Duplicate removal and data type optimization
  - Outlier detection using IQR method with smart removal
  - Business rule validation and consistency checks
  - Data quality metrics reporting

### 2.0 ANALYZE CSV DATA
- **Input**: Cleaned data
- **Output**: Comprehensive data analysis
- **Processing**:
  - Identify column types (numeric, categorical, datetime)
  - Calculate statistical summaries with outlier analysis
  - Detect correlations and strong relationships
  - Analyze data patterns and distributions
  - Generate data quality assessment
  - Create cleaning summary report

### 3.0 GENERATE AI INSIGHTS
- **Input**: Data analysis results
- **Output**: Enhanced structured presentation insights
- **Processing**:
  - Build comprehensive data summary with cleaning details
  - Create enhanced AI prompt with 5-8 bullet points per slide
  - Call OpenAI GPT-3.5-turbo API with detailed requirements
  - Parse JSON response with enhanced slide content
  - Validate and enhance with smart defaults
  - Generate multiple detailed slides (Executive Summary, Key Findings, Quality Assessment)

### 4.0 CREATE CHARTS
- **Input**: Chart specifications from insights
- **Output**: Chart image files
- **Processing**:
  - Generate bar, pie, line, scatter, and heatmap charts
  - Apply professional styling with seaborn
  - Handle proper y-column specifications for all chart types
  - Save as high-resolution PNG files
  - Handle chart creation errors with fallbacks

### 5.0 BUILD PRESENTATION
- **Input**: Enhanced insights and chart files
- **Output**: PowerPoint presentation file
- **Processing**:
  - Create presentation with blank layouts for full control
  - Add title, executive summary, findings, and chart slides
  - Apply consistent formatting with proper spacing
  - Embed chart images with optimal positioning
  - Generate comprehensive bullet points (5-8 per slide)
  - Save as PPTX file with enhanced content

---

## System Interfaces

### External APIs:
- **OpenAI API**: GPT-3.5-turbo for insight generation
- **File System**: CSV input, PPTX output, temporary chart storage

### Internal Libraries:
- **pandas/numpy**: Data processing and analysis
- **matplotlib/seaborn**: Chart generation
- **python-pptx**: PowerPoint file creation
- **openai**: API communication

### Configuration:
- **Environment Variables**: OpenAI API key via .env file
- **Default Settings**: Chart styling, slide dimensions, fonts

This DFD provides a comprehensive view of how data flows through the CSV-to-PowerPoint AI Analyzer system, from initial CSV input to final presentation output.
