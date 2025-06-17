# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Quick Start Commands

**Run the application:**
```bash
python run.py
```

**Install dependencies:**
```bash
pip install pandas numpy openpyxl matplotlib reportlab kivy PyMuPDF
```

## Application Architecture

This is an ASTM Test Calculator for processing acoustic test data. The application follows a **service-oriented architecture** with clear separation of concerns:

### Service Layer

**Calculation Service (`src/core/calculation_service.py`):**
- Unified interface for all test calculations (AIIC, ASTC, NIC, DTC)
- Extracted from main window to separate business logic from UI logic
- Handles test-specific mathematical calculations and ASTM compliance

**Validation Service (`src/core/validation_service.py`):**
- Path validation, file format validation, and data integrity checks
- Provides helpful error suggestions and validation decorators
- Centralized validation logic used throughout the application

**Plotting Service (`src/gui/plotting_service.py`):**
- Separated plot generation from display logic
- Creates matplotlib plots for different test types with ASTM reference contours
- Handles temporary file management and cleanup

### UI Components

**Main Window (`src/gui/main_window.py`):** - **REFACTORED FROM 2,724 → 607 LINES (78% REDUCTION)**
- Focuses on UI coordination and event handling
- Delegates business logic to services
- Uses component-based UI architecture

**UI Components (`src/gui/components/`):**
- `file_input_panel.py` - File path inputs and file selection dialogs
- `test_control_panel.py` - Test action buttons and debug controls  
- `status_panel.py` - Status reporting and error display popups

**Existing Components:**
- `test_plan_input.py` - Form interface for creating/editing test plans
- `analysis_dashboard.py` - Grid-based results visualization dashboard
- `test_plan_manager.py` - Multi-tab test plan management interface

### Core Processing (`src/core/`)

- `test_data_manager.py` - Central coordinator for data loading and file management
- `test_processor.py` - Raw SLM Excel file processing and test data object creation
- `data_processor.py` - ASTM-compliant mathematical calculations and data formatting

### Reports (`src/reports/`)

- `test_data_exporter.py` - CSV report generation and matplotlib plot creation

### Test Types Supported
- **AIIC** - Apparent Impact Insulation Class (floor/ceiling impact sound)
- **ASTC** - Apparent Sound Transmission Class (airborne transmission)
- **NIC** - Noise Isolation Class (background noise isolation)
- **DTC** - Door Transmission Class (partial implementation)

### Data Flow Architecture
1. **Input**: Test plan CSV + SLM data directories
2. **Loading**: TestDataManager validates paths and loads configurations
3. **Processing**: TestProcessor extracts measurement data → test objects
4. **Calculation**: ASTM calculations performed on measurement data
5. **Storage**: Results stored in `test_data_collection` dictionary
6. **Visualization**: Matplotlib plots with reference contours
7. **Export**: CSV reports with accompanying plots

### Key Data Structures
- `RoomProperties` dataclass - Test environment specifications
- Test-specific data classes: `AIICTestData`, `ASTCTestData`, `NICTestData`, `DTCtestData`
- Central storage: `test_data_collection` dictionary in TestDataManager

## Development Guidelines

**Code Style:**
- Follow PEP 8 style guidelines
- Use descriptive variable names reflecting data content
- Prefer vectorized operations over explicit loops
- Use method chaining for data transformations when possible

**Data Processing:**
- Use pandas for data manipulation and analysis
- Implement data quality checks at analysis beginning
- Handle missing data appropriately (imputation, removal, flagging)
- Use try-except blocks for error-prone operations

**File Structure Requirements:**
- Test plan CSV format is critical - errors will result if format not followed
- Maintain `/Exampledata` directory structure for proper functionality
- SLM data must be in Excel format with specific naming conventions

## Important File Locations

**Example Data:** `/Exampledata/` contains properly formatted test plans and raw data
**Backups:** `/backups/` contains automatic backup storage
**Entry Point:** `run.py` - Application entry point that sets up Python path

## Development Status

This application is in "Minimum Lovable Prototype" phase. No formal testing framework, build configuration, or CI/CD is currently implemented.