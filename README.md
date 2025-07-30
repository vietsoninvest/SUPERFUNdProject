# Australian Superfund Analysis and Visualization

## Contents

- [Project Overview](#project-overview)  
- [Project Goals](#project-goals)  
- [Expected Outcomes](#expected-outcomes)  
- [Tools and Techstack (To Be Updated)](#tools-and-techstack-to-be-updated)  
- [Development Scope and Timeline](#development-scope-and-timeline)  
  - [Sprint 1: PHD Data Collection, Standardization, and Analysis (2–3 weeks)](#sprint-1-phd-data-collection-standardization-and-analysis-2–3-weeks)  
  - [Sprint 2: PHD Visualization (3–4 weeks)](#sprint-2-phd-visualization-3–4-weeks)  
  - [Sprint 3: UP Data Collection, Standardization, and Analysis](#sprint-3-up-data-collection-standardization-and-analysis)  
  - [Sprint 4: Performance Visualization & Combined Insights](#sprint-4-performance-visualization--combined-insights)  
  - [Sprint 5: Testing and Final Deployment](#sprint-5-testing-and-final-deployment)  
  - [Sprint 6: Documentation Website Development](#sprint-6-documentation-website-development)  
- [Deliverables](#deliverables)  
- [Obstacles and Challenges](#obstacles-and-challenges)  

---

## Project Overview

This project aims to provide an interactive and data-driven platform for understanding and comparing Australian superannuation funds based on their **Portfolio Holding Disclosure (PHD)** data and **Unit Pricing (UP)**.  
By visualizing investment types, asset allocations, and performance metrics, the project will help users — especially individual investors — make informed decisions when choosing a super fund.

---

## Project Goals

- Build a robust pipeline to collect, clean, and standardize PHD and Unit Pricing data across multiple Australian super funds.  
- Analyze asset allocation, sector exposure, and investment strategies of each fund.  
- Compare fund asset allocation and performance through historical unit pricing.  
- Visualize the insights using Power BI dashboards hosted on a report server.  
- Build a static website to host documentation, insights, and provide access to dashboards.  

---

## Expected Outcomes

A set of **Power BI dashboards** enabling users to explore:

- Asset class allocations (e.g., cash, property, foreign currency, etc.)  
- Industry-specific investment exposure (e.g., tech, mining, healthcare, etc.)  
- Geographic allocation and concentration risks  
- Unit price-based performance over time  

A lightweight **static website** to:

- Explain data methodology  
- Host documentation and updates  
- Link or embed Power BI reports  

Tools and documentation that assist users in choosing the right super fund based on real data.

---

## Tools and Techstack (To Be Updated)

- **Data Collection & Cleaning**: Python (Pandas, BeautifulSoup, Requests), Power Query  
- **Data Storage**: CSV / MySQL Server / Google Cloud Platform  
- **Visualization**: Power BI  
- **Hosting**: Power BI Report Server / Power BI Service / SQL Server Reporting Services (SSRS)  
- **Static Website**: GitHub Pages  
- **Version Control & Documentation**: Git, GitHub  

---

## Development Scope and Timeline

### Sprint 1: PHD Data Collection, Standardization, and Analysis (2–3 weeks)

- Identify and gather PHD datasets from super fund sources  
- Build Python scripts to ingest and standardize the data (fund name, date, asset type, sector, etc.)  
- Perform Exploratory Data Analysis (EDA):  
  - Dominant sectors  
  - Asset class distribution  
  - Compare across funds  
- Deliver insights  

### Sprint 2: PHD Visualization (3–4 weeks)

- Create Power BI dashboards showing:  
  - Asset class allocation (equity, fixed income, etc.)  
  - Industry-level exposure  
  - Regional/geographic breakdown  
  - Interactive fund selection and filtering  
- Research deployment of visualizations to Power BI Report Server or embedding in the static site  

### Sprint 3: UP Data Collection, Standardization, and Analysis

- Use web scraping techniques to collect historical unit pricing data for each fund  
- Standardize data formats (e.g., daily NAV, fund ID, date)  
- Analyze performance metrics:  
  - Growth trends  
  - Cumulative return  
  - Performances by sector  
  - Volatility and price movements  

### Sprint 4: Performance Visualization & Combined Insights

- Build Power BI dashboards comparing:  
  - Fund performances over time  
  - Growth comparisons  
  - Correlations or divergence between funds  
- Integrate with PHD dashboard  
- Successfully deploy Power BI dashboards to a selected server  

### Sprint 5: Testing and Final Deployment

- Validate data accuracy and visual logic across all dashboards  
- Write a data refresh and update guide  
- Final documentation for website and GitHub repo  

### Sprint 6: Documentation Website Development

- Build a clean, readable static website for hosting:  
  - Dashboard links  
  - Documentation  
  - FAQs  
  - Changelog  

---

## Deliverables

- Standardized and documented datasets (PHD + Unit Pricing)  
- Power BI dashboards with:  
  - Investment breakdown  
  - Industry and sector heatmaps  
  - Fund performance comparisons  
- A static website hosting:  
  - Dashboard links  
  - Documentation  
  - FAQs and changelog  

---

## Obstacles and Challenges

_(Add details here if applicable. Examples may include data inconsistency, limited availability of UP data, deployment issues, etc.)_
