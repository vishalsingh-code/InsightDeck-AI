# Visual Data Flow Diagrams - CSV-to-PowerPoint AI Analyzer

## Context Diagram (Level 0) - System Overview

```
                     ┌──────────────────────────────────────────────┐
                     │                                              │
         [User]      │         CSV-to-PowerPoint                    │      [User]
            │        │           AI Analyzer                        │         │
            │        │                                              │         │
         CSV File ───┼─────────────────────────────────────────────┼───► PowerPoint
            │        │                                              │    Presentation
            │        │              ┌─────────────┐                 │         │
            └────────┼─────────────►│   System    │◄────────────────┼─────────┘
                     │              │   Process   │                 │
      [OpenAI API]   │              └─────────────┘                 │   [File System]
            │        │                     │                        │         │
         AI Response ┼─────────────────────┘                        │         │
            │        │                                              │         │
            └────────┼──────────────────────────────────────────────┼───► Temporary
                     │                                              │    Chart Files
                     │                                              │
                     └──────────────────────────────────────────────┘
```

## Level 1 DFD - Main Process Flow

```
                                    CSV File
                                       │
                                       ▼
                           ┌─────────────────────────┐
                           │       Process 1.0       │
                           │    LOAD & ANALYZE       │◄──── Data Quality
                           │        CSV              │      Rules
                           └───────────┬─────────────┘
                                       │
                               Analysis Results
                                       │
                                       ▼
                           ┌─────────────────────────┐
                           │       Process 2.0       │────┐
                           │     GENERATE AI         │    │ API Request
                           │      INSIGHTS           │◄───┘ (OpenAI)
                           └───────────┬─────────────┘
                                       │
                            Structured Insights
                                       │
                                       ▼
                           ┌─────────────────────────┐
                           │       Process 3.0       │
                           │      CREATE             │
                           │      CHARTS             │
                           └───────────┬─────────────┘
                                       │
                               Chart Image Files
                                       │
                                       ▼
                           ┌─────────────────────────┐
                           │       Process 4.0       │◄──── Chart Files
                           │       BUILD             │
                           │    PRESENTATION         │
                           └───────────┬─────────────┘
                                       │
                                       ▼
                             PowerPoint Presentation
```

## Level 2 DFD - Detailed Process Breakdown

### Process 1.0 - CSV Analysis Pipeline

```
    CSV File
       │
       ▼
┌─────────────┐    Raw Data     ┌─────────────┐
│  Process    │───────────────►│  D1: RAW    │
│   1.1       │                │    DATA     │
│ LOAD CSV    │                └─────────────┘
└─────────────┘                       │
       │                              │
       │ Processing                   │ Read
       │ Control                      │ Access
       │                              ▼
       │                    ┌─────────────────┐
       │                    │   Process 1.2   │
       │                    │ BASIC ANALYSIS  │
       │                    │ • Data Types    │
       │                    │ • Missing Vals  │
       │                    │ • Basic Stats   │
       │                    └─────────┬───────┘
       │                              │
       │                     Basic Statistics
       │                              │
       │                              ▼
       │                    ┌─────────────────┐
       │                    │   Process 1.3   │
       │                    │ADVANCED ANALYSIS│
       │                    │ • Correlations  │
       │                    │ • Outliers      │
       │                    │ • Patterns      │
       │                    └─────────┬───────┘
       │                              │
       │                    Comprehensive Analysis
       │                              │
       │                              ▼
       │                    ┌─────────────────┐
       └───────────────────►│ D2: ANALYSIS    │
                            │    RESULTS      │
                            │ • Stats Summary │
                            │ • Correlations  │
                            │ • Data Quality  │
                            │ • Insights      │
                            └─────────────────┘
```

### Process 2.0 - AI Insights Generation

```
Analysis Results
       │
       ▼
┌─────────────────┐   Formatted      ┌─────────────────┐
│   Process 2.1   │   Data Summary   │   Process 2.2   │
│  BUILD DATA     │─────────────────►│   CALL OPENAI   │
│   SUMMARY       │                  │      API        │
└─────────────────┘                  └─────────┬───────┘
       │                                       │
       │ Context Data                    AI Response (JSON)
       │                                       │
       │                                       ▼
       │                             ┌─────────────────┐
       │                             │   Process 2.3   │
       │                             │  PARSE & CLEAN  │
       │                             │   AI RESPONSE   │
       │                             └─────────┬───────┘
       │                                       │
       │                              Parsed Insights
       │                                       │
       │                                       ▼
       │                             ┌─────────────────┐
       │                             │   Process 2.4   │
       │                             │   VALIDATE &    │
       │                             │    ENHANCE      │
       │                             └─────────┬───────┘
       │                                       │
       │                            Enhanced Insights
       │                                       │
       │                                       ▼
       └────────────────────────────►┌─────────────────┐
                                     │ D3: STRUCTURED  │
                                     │    INSIGHTS     │
                                     │ • Title         │
                                     │ • Key Findings  │
                                     │ • Chart Specs   │
                                     │ • Slide Content │
                                     └─────────────────┘
```

### Process 3.0 - Chart Creation Pipeline

```
Structured Insights + Analysis Results
                │
                ▼
        ┌─────────────────┐
        │   Process 3.1   │
        │ SELECT OPTIMAL  │
        │   CHART TYPES   │
        │ • Bar Charts    │
        │ • Pie Charts    │
        │ • Line Charts   │
        │ • Scatter Plots │
        │ • Heatmaps      │
        └─────────┬───────┘
                  │
           Chart Specifications
                  │
                  ▼
        ┌─────────────────┐    Prepared Data   ┌─────────────────┐
        │   Process 3.2   │──────────────────►│   Process 3.3   │
        │  PREPARE DATA   │                   │ RENDER CHARTS   │
        │ • Filter Data   │                   │ • Matplotlib    │
        │ • Aggregate     │                   │ • Seaborn       │
        │ • Transform     │                   │ • Styling       │
        └─────────────────┘                   └─────────┬───────┘
                                                        │
                                               Chart Images (PNG)
                                                        │
                                                        ▼
                                              ┌─────────────────┐
                                              │  D4: CHART      │
                                              │     FILES       │
                                              │ • chart_0.png   │
                                              │ • chart_1.png   │
                                              │ • chart_2.png   │
                                              │ • ...           │
                                              └─────────────────┘
```

### Process 4.0 - Presentation Building

```
Structured Insights + Chart Files
                │
                ▼
        ┌─────────────────────┐
        │    Process 4.1      │
        │  CREATE SLIDE       │
        │    STRUCTURE        │
        │ • Title Slide       │
        │ • Content Slides    │
        │ • Chart Slides      │
        └─────────┬───────────┘
                  │
            Slide Specifications
                  │
                  ▼
        ┌─────────────────────┐    Slide Data    ┌─────────────────────┐
        │    Process 4.2      │─────────────────►│    Process 4.3      │
        │   CREATE SLIDES     │                  │   EMBED CHARTS      │
        │ • Text Content      │                  │ • Image Placement   │
        │ • Formatting        │                  │ • Positioning       │
        │ • Layout            │                  │ • Sizing            │
        └─────────────────────┘                  └─────────┬───────────┘
                                                           │
                                                  Complete Slides
                                                           │
                                                           ▼
                                                 ┌─────────────────────┐
                                                 │    Process 4.4      │
                                                 │  FINALIZE & SAVE    │
                                                 │ • Apply Themes      │
                                                 │ • Cleanup Temp      │
                                                 │ • Save PPTX         │
                                                 └─────────┬───────────┘
                                                           │
                                                           ▼
                                                  PowerPoint File (.pptx)
```

## Data Store Details

```
┌─────────────────────────────────────────────────────────────────────────┐
│                              DATA STORES                               │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  D1: RAW_DATA                    D2: ANALYSIS_RESULTS                   │
│  ┌─────────────────┐            ┌─────────────────────────────────┐     │
│  │ • CSV Content   │            │ • Statistical Summaries         │     │
│  │ • DataFrame     │            │ • Correlation Matrices          │     │
│  │ • Original      │    ───►    │ • Data Quality Metrics         │     │
│  │   Structure     │            │ • Pattern Recognition          │     │
│  │ • Column Names  │            │ • Column Classifications       │     │
│  └─────────────────┘            └─────────────────────────────────┘     │
│           │                                      │                      │
│           │                                      │                      │
│           ▼                                      ▼                      │
│  D4: CHART_FILES                 D3: STRUCTURED_INSIGHTS                │
│  ┌─────────────────┐            ┌─────────────────────────────────┐     │
│  │ • PNG Images    │            │ • Presentation Title            │     │
│  │ • Temporary     │            │ • Key Insights & Findings      │     │
│  │   Storage       │    ◄───    │ • Chart Recommendations        │     │
│  │ • High DPI      │            │ • Slide Content Structure      │     │
│  │ • Multiple      │            │ • AI-Generated Content         │     │
│  │   Chart Types   │            │ • Smart Fallback Data          │     │
│  └─────────────────┘            └─────────────────────────────────┘     │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

## System Data Flow Summary

```
INPUT                 PROCESSING STAGES              OUTPUT
──────               ──────────────────             ──────

┌─────────┐          ┌─────────────────┐          ┌─────────────┐
│ CSV     │          │  1. Data        │          │ PowerPoint  │
│ File    │─────────►│     Loading     │─────────►│ Presentation│
│         │          │  2. Analysis    │          │ (.pptx)     │
└─────────┘          │  3. AI Insights │          └─────────────┘
                     │  4. Chart       │
┌─────────┐          │     Creation    │          ┌─────────────┐
│ OpenAI  │          │  5. Slide       │          │ Temporary   │
│ API     │◄────────►│     Building    │─────────►│ Chart Files │
│         │          └─────────────────┘          │ (Cleanup)   │
└─────────┘                                       └─────────────┘

EXTERNAL ENTITIES:                    INTERNAL PROCESSES:
• User (Input/Output)                • Data Analysis Engine
• OpenAI API (AI Service)           • Chart Generation System  
• File System (Storage)             • Presentation Builder
```

This comprehensive DFD documentation shows how data flows through the CSV-to-PowerPoint AI Analyzer system at multiple levels of detail, from high-level context to detailed process breakdowns.
