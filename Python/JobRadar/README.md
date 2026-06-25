# JobRadar



## Overview

JobRadar is a Python application that connects to Gmail and automatically analyzes job alerts from LinkedIn and Google Job Alerts.

The tool extracts job opportunities, scores them against my background, categorizes them by fit, and exports organized Excel reports to support a structured job search process.

## Purpose 
Purpose

The goal of JobRadar is to reduce the time spent manually reviewing job alerts and help prioritize the best opportunities faster.
## Business Problem

Searching for jobs manually across multiple platforms is time consuming and makes it difficult to prioritize opportunities.

JobRadar reduces manual effort by collecting job opportunities from Gmail and ranking them based on career goals, experience, and target roles.

## Who This Is For

JobRadar is designed for job seekers who receive large numbers of job alerts and want to spend less time reviewing emails and more time applying to the right opportunities.


## Why I Built This

As someone transitioning into Salesforce Administration, Business Analysis, and AI-driven workflow automation, I wanted to build a tool that solved a real problem I faced every day.

Instead of manually reviewing hundreds of job alerts, I wanted a system that could automatically organize opportunities, prioritize the best matches, and reduce the time spent searching so I could spend more time preparing for interviews and building new skills.

JobRadar began as a personal productivity tool and has evolved into a portfolio project demonstrating Python development, API integration, automation, data analysis, and product thinking.

## Current Features

### Gmail Integration

* Connects to Gmail using OAuth
* Reads job alert emails
* Supports LinkedIn Job Alerts
* Supports Google Job Alerts

### Job Extraction

* Extracts job title
* Extracts company
* Extracts location
* Captures email source
* Captures email subject
* Captures email date

### Job Scoring Engine

Jobs are scored based on:

* Salesforce Administrator keywords
* CRM experience
* Business Analyst keywords
* Operations Analyst keywords
* RevOps and Sales Operations keywords
* Remote opportunities
* Hybrid opportunities
* Minnesota opportunities

### Resume Recommendation Engine

Suggests the best resume version:

* Salesforce Resume
* Business Analyst Resume
* Operations / Supply Chain Resume
* Sales Ops / RevOps Resume

### Job Categorization

Classifies jobs into:

* Excellent Match
* Strong Match
* Possible Match
* Low Match

### Application Prioritization

Creates recommendations:

* Apply Now
* High Priority
* Review
* Skip

### Recruiter Lead Detection

Identifies:

* LinkedIn invitations
* Recruiter outreach
* Hiring manager messages

### Reporting

Generates:

* job_radar_results.xlsx
* apply_now.xlsx
* high_priority.xlsx
* review.xlsx
* skip.xlsx
* weekend_apply_list.xlsx
* recruiter_leads.xlsx

### Dashboard

Provides:

* Total Jobs
* Apply Now Count
* High Priority Count
* Review Count
* Skip Count
* Weekend Apply List Count
* Recruiter Lead Count

## Technology Stack
## Technology Stack

* Python 3
* Gmail API
* Google OAuth 2.0
* Pandas
* OpenPyXL
* BeautifulSoup
* CSV Processing
* Excel Automation
* Visual Studio Code
* Git
* GitHub

## Future Enhancements

## Future Enhancements

### V1.7
- HTML cleanup
- Better duplicate detection
- Recruiter message cleanup

### V1.8
- Persistent application tracking
- Company watchlist

### V1.9
- Recruiter CRM
- Networking tracker

### V2.0
- AI job fit summaries
- AI resume tailoring
- AI cover letter generation
- AI interview preparation
- AI career recommendations

### V3.0
- Desktop application (.exe)
- Personal job search dashboard
- One-click apply workflow
- AI Career Copilot integration




## Example Workflow

1. Connect to Gmail using OAuth.
2. Read LinkedIn and Google Job Alert emails.
3. Extract job information.
4. Score each opportunity.
5. Recommend the best resume.
6. Categorize the opportunity.
7. Generate Excel reports.
8. Prioritize which jobs to apply for first.


## Version History

### V1.0
- Gmail API connection

### V1.1
- LinkedIn Job Alert extraction

### V1.2
- Google Job Alert support

### V1.3
- Job scoring engine

### V1.4
- Resume recommendation engine

### V1.5
- Recruiter lead detection

### V1.6
- Excel dashboard
- Weekend Apply List
- Company scoring
- Location filtering
- Application prioritization


