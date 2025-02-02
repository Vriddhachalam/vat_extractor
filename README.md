# Vat Extractor

## Description

A brief description of your project goes here. Explain what your project does and any other important information.

---

## Prerequisites

Before running the project, make sure you have the following installed:

- [Python](https://www.python.org/downloads/) (version 3.12)

---

## Installation

### 1. Clone the repository

Clone the project repository to your local machine using the following command:

```bash
git clone https://github.com/Vriddhachalam/vat_extractor.git
```

### 2. Placing the invoices (PDF files)

Place files under inward and outward directories inside invoices directory of program

```
vat_extractor/program/invoices/inward/
```
```
vat_extractor/program/invoices/outward/
```
### 3. Run the executable (Windows only)

Once you’ve cloned the repo, go to the project directory:

```bash
cd vat_extractor/program/
```
The run

```bash
./vat_extractor.exe
```
or run it manually from your ui

## Modify the logic (Advanced)

### 1. Virtual Environment (Optional but recommended)

It's a good practice to use a virtual environment to avoid conflicts with system-wide packages. To create a virtual environment, run:

```bash
python -m venv venv
```

Activate the virtual environment. The activation command may vary based on your operating system:

- **Windows**:

```bash
  source venv/Scripts/activate
```

- **Linux**:

```bash
  source venv/bin/activate
```

### 2. Install dependencies from requirements.txt

Once the virtual environment is active, install all necessary dependencies listed in requirements.txt:

```bash
pip install -r requirements.txt
```

### 3. Copy the file vat_extractor.py to program directory

Copy the code to program folder
Run the following command when in root path of repository
```bash
cp vat_extractor.py program/
```

### 4. Modify the code to your needs 
Then run when in program/
```bash
python vat-extractor.py
```

### 5. Recompile the code to EXE:
Install pyinstaller

```bash
pip install pyinstaller
```
Then run 

```bash
pyinstaller --onefile vat_extractor.py
```
After successfully running:

```bash
cd build
```
You have the vat_extractor.exe

ENJOY!


