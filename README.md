# SOX-Projects

In my project, I’ve developed a set of macros to streamline tasks and improve productivity. These macros serve different purposes:

## Macros Overview

### "Check Resolution" Macro
- **Purpose**: Verifies if the entire content is visible in a screenshot.
- **How it works**: Extracts the width and height from cells B7/B8 and captures a screenshot.
- **Benefit**: Ensures that critical information is captured accurately.

### "Get Coordinates" Macro
- **Purpose**: Retrieves coordinates for a specific element (the small “triangle” in SAP).
- **Workflow**: When activated, you have 4 seconds to position your mouse over the “triangle.”
- **Result**: Displays a message box with the precise coordinates (xpos/ypos), which are then automatically recorded in cells B4/B5.
- **Importance**: Essential for capturing detailed project status screenshots.

### "Check Coordinates" Macro
- **Purpose**: Validates the correctness of the triangle’s coordinates.
- **Functionality**: Verifies if the recorded coordinates align with the expected position.
- **Significance**: Ensures accurate interaction with the “triangle” during subsequent screenshot captures.

## Main Project Macros

### "POC" Macro Overview (similar Macro for CCM, CCM Service, WAR)
- **Worksheet**: The "POC" worksheet serves as the foundation.
- **Column A**: Contains project numbers for which we need status information.
- **Column B**: Represents the "Nodes" (locations within the SAP code) associated with these projects.
- **Column C**: Holds the select command relevant to each project.
- **Functionality**: The macro processes this data to retrieve project status efficiently.

### "Cleaning Sheets" Macro
- **Purpose**: Automates the process of deleting screens from all worksheets before the next Sox control.
- **Workflow**:
  - When activated, it promptly removes screens without manual intervention.
  - Ensures a clean slate for the upcoming session.
- **Benefits**:
  - Saves time by eliminating the need for manual deletion.
  - Enhances consistency across all worksheets.
  - Facilitates a smooth transition to the next control.

---

Transitioning from an hour of manual work to a mere 5 minutes using these project macros represents a remarkable improvement. By automating repetitive tasks, we not only saved valuable time but also ensured consistency and accuracy in our project documentation.

![image](https://github.com/user-attachments/assets/7a1c559d-379c-4fe8-aa3b-ba7e4b7623eb)

