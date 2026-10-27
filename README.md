# ⚡ Automated Web Data Extraction Pipeline

An end-to-end Python automation tool designed to process multi-country web sources efficiently in **under 5 minutes**. 

This project demonstrates structured data ingestion, configuration-driven web automation, and clean modular Python architecture.

### Tech Stack & Libraries
- **Language:** Python
- **Libraries/Tools:** Pandas, Requests, PyODBC / OpenPyXL


## 📌 Features

- **Configuration-Driven:** Uses an Excel matrix (`master.xlsx`) to dynamically control target websites, country metadata, and task execution without hardcoding parameters.
- **Fast & Efficient:** Optimized workflow that completes multi-site extraction and processing in **less than 5 minutes**.
- **Modular Architecture:** Clear separation between data parsing/configuration management and core execution logic.

---

## 📁 Repository Structure

```text
├── master.xlsx       # Configuration sheet containing URLs, country mappings, and parameters
├── main.py           # Entry point: Reads master.xlsx and orchestrates task execution
├── sourcecode.py     # Core automation scripts, parsing logic, and execution functions
└── README.md         # Project documentation
