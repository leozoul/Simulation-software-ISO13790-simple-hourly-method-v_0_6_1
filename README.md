# Simulation Software – ISO 13790 Simple Hourly Method (v0.6.1)

This repository contains a **building energy simulation software** based on a **modified version of the ISO 13790:2017 Simple Hourly Method**.

---

## Overview

The software consists of **two main scripts** that work together:

1. **Building Structure & Calculation Engine**  
   Implements the building model and calculation procedure according to the **ISO 13790 Simple Hourly Method**.

2. **Graphical User Interface (GUI)**  
   Allows users to:
   - Define and structure a building
   - Import climate files
   - Configure user preferences  
   - Perform **dynamic daily energy simulations**

   These simulations calculate the **total energy demand of the modeled building**.

---

## Features

- The building is modeled as a **single lumped thermal zone** interacting with the external environment.
- Currently supports simulation of:
  - **Residential buildings**
  - **Office buildings**
- Buildings can be **saved and loaded** using a specially formatted **`.xlsx` (Excel) file**.

---

## Required Python Modules

Install the following Python packages:

- **PySimpleGUI** (version 4)
- **Matplotlib**
- **NumPy**
- **Pandas**

You can install them using:

```bash
pip install PySimpleGUI matplotlib numpy pandas
```

---

## Usage

1. **Download or clone** all repository files.
2. **Install the required Python modules**.
3. Run the graphical interface:

```bash
python ISO13790shm_graphical_environment_v0_6_1.py
```

4. Follow the instructions in **`Documentation.pdf`**.

---

## Author

**Leonidas Zouloumis**

📧 Email: leozoul@gmail.com  
🔗 LinkedIn:  
https://www.linkedin.com/in/leonidas-zouloumis-838467100/
