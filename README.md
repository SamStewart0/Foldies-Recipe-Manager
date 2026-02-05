# Foldies-Recipe-Manager
Excel sheet created for Foldies MCR to manage company recipies

Cooking Company: Recipe Management & Automation Tool

Project Overview
This project was developed as a solution for a restaurant I worked at to automate their recipe process. The tool transforms a manual data entry task into a dynamic system, ensuring that new recipes are standardized, indexed, and easily accessible.

The Problem
The company was manually copying sheets and updating a central list whenever a new recipe was created. This was time-consuming and prone to broken links and inconsistent formatting.

I developed a VBA-driven macro that:
Dynamic Templating: clones a standardized "ADD_RECIPES" template to create a new, dedicated recipe sheet.
Auto-Indexing: updates a central "INDEX" sheet with a clickable hyperlink to the new recipe, ensuring the workbook stays organized.
State Management: specific input ranges in the template after a successful save, preparing the UI for the next entry without deleting vital formulas.
Robust Error Handling: defensive programming to check for missing sheets or empty names, preventing the workbook from crashing.

How to Use
1.  Open `Recipe_Manager.xlsm`.
2.  Navigate to the ADD_RECIPES sheet.
3.  Enter the recipe name in cell C4 and fill in the details.
4.  Run the `AddNewRecipe` macro (assigned to a button).
5.  Check the INDEX sheet to see your new clickable entry!
