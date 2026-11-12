# Hospital Food Consumption & Nutritional Analysis

### University of Southampton | Software and Hardware Research Assistant

**Project Area:** Computer Vision, Data Analysis, Nutritional Science, Python Development

**Duration:** 10-week Research Internship

## Project Overview

This repository contains a selection of Python scripts and screenshots demonstrating the work I completed during my internship as a Software and Hardware Research Assistant at the University of Southampton.

During this 10-week internship, I contributed to an interdisciplinary research project within the Life Sciences department, focusing on nutritional science and the monitoring of food consumption in hospital patients.

The overarching aim of the research project was to develop a device capable of estimating the nutritional intake of patients by analysing images of their meals taken before and after consumption. By identifying the food served and estimating how much remains, the system aims to calculate the nutritional content of the food consumed.

This approach has the potential to support healthcare professionals in monitoring patients' nutritional intake while reducing the need for manual food consumption assessments.

## Initial Dataset

Upon joining the project, I was provided with a dataset containing **717 hospital meals**, each with a pair of before-and-after images alongside accompanying meal information stored in a spreadsheet.

The data had been collected using the second iteration of the research device during a 20-week study on a hospital stroke ward, covering approximately 50 patients.

The accompanying spreadsheet contained:

- Anonymised patient identifiers and unique meal identifiers.
- The date, time and study week associated with each meal.
- Descriptions of the foods served.
- The recorded serving weights of individual foods.
- Weight measurements captured by the device before and after each meal.

My primary responsibility was to develop software capable of processing this dataset and estimating how much of each food item had been consumed, allowing the results to be combined with nutritional composition data.

## Project Objectives

The main objectives of my work were to:

1. Develop a reliable method of pairing and processing before-and-after meal images.
2. Investigate image segmentation techniques to distinguish food from plates, trays and surrounding objects.
3. Estimate the percentage of each food item consumed using the available image and weight data.
4. Implement quality-control procedures to identify unreliable image analysis results.
5. Combine food consumption estimates with nutritional composition data to estimate nutritional intake.
6. Develop an accessible interface through which researchers could inspect meal images and view consumption and nutritional results.

## Image Processing and Food Segmentation

One of the main technical challenges was identifying the food present in each image and estimating how much remained after consumption.

The dataset contained meals with different plate arrangements, food types, lighting conditions and image quality. Consequently, a single segmentation approach was not suitable for every meal.

I investigated and implemented image-processing methods to identify food regions and compare the before-and-after images.

These included:

- **Image pairing:** Organising the dataset and associating before-and-after images with the correct meal identifiers.
- **Image segmentation:** Generating masks to isolate relevant food regions from the background.
- **Region-based analysis:** Using different approaches to account for the arrangement of food on plates.
- **Quality control:** Identifying problematic masks and meals for which image-based consumption estimates could not be considered reliable.

The resulting pipeline was designed to process large numbers of meals while allowing individual results to be reviewed where necessary.

## Food Consumption Estimation

Once the image segmentation results had been generated, the next stage involved estimating the proportion of each food that had been consumed.

For foods with usable before-and-after measurements, the consumption percentage was estimated using:

**Percentage Consumed = ((Before Quantity − After Quantity) / Before Quantity) × 100**

These percentages were then combined with the recorded serving weights to estimate the amount of food consumed in grams.

### Handling Unreliable Images

An important aspect of the project was developing methods to handle meals where image analysis alone was insufficient.

Some meals contained missing foods, unreliable segmentation masks or inconsistent image measurements.

To address this, I implemented additional processing logic, including:

- Quality-control classifications for unreliable segmentation results.
- Alternative consumption estimates based on recorded weight differences.
- Handling of meals with zero consumption.
- Reference-based estimation methods for certain foods with missing image-derived measurements.

For meals where image-based estimates were unsuitable but valid weight measurements were available, a weight-based estimate could be calculated:

**Percentage Consumed = (Measured Weight Difference / Total Served Weight) × 100**

This allowed additional meals to be included in the analysis rather than automatically discarding them because of poor image quality.

## Nutritional Analysis

Following the consumption estimation stage, I developed a workflow to calculate the nutritional content of the food consumed.

This involved integrating the food consumption results with the **McCance and Widdowson Composition of Foods Integrated Dataset (2021)**.

Each food item was associated with a corresponding nutritional composition code, allowing the nutritional content to be calculated from the estimated consumed weight.

A particular challenge was handling composite meals containing multiple ingredients.

For example, a prepared meal could be separated into its constituent ingredients using predefined recipe proportions. The estimated consumed weight could then be distributed between those ingredients and used to calculate their respective nutritional contributions.

This approach enabled the system to estimate nutritional intake for individual foods and entire meals.

## Streamlit Application

To make the results easier to access and interpret, I developed an interactive application using **Streamlit**.

The application allows researchers to search for meals using patient identifiers and inspect the associated data without manually navigating the underlying files.

Its functionality includes:

- Searching for meal records using patient identifiers.
- Displaying before-and-after meal images.
- Presenting estimated food consumption percentages.
- Displaying estimated consumed food weights.
- Calculating and presenting nutritional intake.
- Highlighting cases where consumption estimates are unavailable or require further review.

The application provides a more accessible way of examining the dataset and reviewing the outputs of the analysis pipeline.

## Technical Challenges and Problem Solving

Several challenges emerged throughout the project, requiring iterative development and debugging.

**1. Inconsistent meal image arrangements**

Some meals contained multiple plates or different arrangements between the before-and-after images. This complicated the process of comparing food regions and required additional validation of image pairings.

**2. Unreliable segmentation results**

Image masks were not consistently accurate across every meal. Quality-control checks and alternative estimation methods were introduced to reduce the reliance on unreliable results.

**3. Complex food descriptions**

Meal descriptions frequently contained multiple foods, composite dishes and inconsistent formatting. Parsing and standardisation were required to link these descriptions to nutritional composition records.

**4. Integrating multiple data sources**

The analysis required information from image files, weight measurements, meal descriptions, food composition datasets and manually reviewed quality-control results. These sources needed to be combined into a consistent processing pipeline.

**5. Developing an accessible interface**

The final results needed to be understandable to researchers without requiring them to interact directly with Python scripts or large spreadsheets.

## Technologies Used

| Technology | Application |
|---|---|
| Python | Main development language |
| pandas | Data cleaning, processing and merging |
| NumPy | Numerical operations and image data processing |
| OpenCV | Image processing and segmentation |
| Matplotlib | Visualisation and debugging |
| Streamlit | Interactive application development |
| Excel / openpyxl | Structured input and output data |
| McCance and Widdowson Dataset | Nutritional composition reference |

## Repository Contents

This repository contains selected Python scripts and screenshots illustrating the developed analysis pipeline and user interface.

The full hospital dataset is not included due to data confidentiality considerations.

## Skills Demonstrated

Through this project, I gained practical experience in:

- Developing Python-based data processing pipelines.
- Applying computer vision techniques to real-world image datasets.
- Cleaning, validating and integrating data from multiple sources.
- Designing fallback methods for incomplete or unreliable data.
- Working with nutritional composition datasets.
- Developing interactive applications using Streamlit.
- Debugging complex workflows and validating calculated outputs.
- Communicating technical findings within an interdisciplinary research environment.

## Project Outcomes

The project resulted in a Python-based workflow capable of processing hospital meal data, estimating food consumption, and calculating associated nutritional intake.

The Streamlit application provided a practical interface for inspecting meal records, reviewing image-analysis results and accessing nutritional estimates.

My contribution focused on developing and integrating the software required to transform the collected meal images and accompanying data into usable consumption and nutritional information.

The work formed part of a wider research effort to improve nutritional monitoring in healthcare environments.
