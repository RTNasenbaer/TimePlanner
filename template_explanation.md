# DOCX Template Usage Guide

## Template File Structure

Your DOCX template should be structured as follows:

### Header Information
The template should contain these placeholders that will be replaced:

- **{{theme}}** - Name of the training plan
- **{{name}}** - Trainer name (from settings)
- **{{playNumber}}** - Number of players (from settings)
- **{{tools}}** - Combined list of all tools from all segments
- **{{time}}** - Total duration of the plan
- **{{requirements}}** - Player requirements (from settings)
- **{{team}}** - Team information (from settings)

### Sections Table
The template should contain a table (preferably the second table in the document) with the following structure:

| Time Slot | Goal | Organisation | Explanation | Tools |
|-----------|------|--------------|-------------|-------|
| (header)  | (header) | (header) | (header) | (header) |

The application will:
1. Clear all existing rows except the header
2. Add one row per segment with:
   - **Time Slot**: Start time - End time (HH:MM format)
   - **Goal**: Segment name
   - **Organisation**: "Spielform" or "Übungsform"
   - **Explanation**: Detailed explanation of the exercise
   - **Tools**: Tools needed for this segment

## How to Use

1. **Configure Settings**: Go to "Einstellungen" → "Einstellungen..." to set:
   - Your trainer name
   - Number of players
   - Player requirements
   - Team information

2. **Create Segments**: Use "Abschnitt hinzufügen" to add segments with:
   - Name
   - Duration
   - Organisation type (Spielform/Übungsform)
   - Detailed explanation
   - Required tools

3. **Export**: Go to "Datei" → "Als DOCX exportieren" (Ctrl+D):
   - Select your template file
   - Choose output location
   - The application will fill in all placeholders and create the table

## Example Template Structure

```
Training Plan: {{theme}}
Trainer: {{name}}
Players: {{playNumber}}
Duration: {{time}}
Requirements: {{requirements}}
Team: {{team}}
Tools needed: {{tools}}

[Table with columns: Time Slot | Goal | Organisation | Explanation | Tools]
```

## Notes

- Make sure your template has exactly the placeholders shown above
- The table should be the second table in your document
- All sections will be automatically sorted by time
- Tools from all segments will be combined in the {{tools}} placeholder
- Time slots will be automatically calculated and formatted as HH:MM