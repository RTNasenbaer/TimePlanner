# Setup Docx file
- Trainingseinheit_Name_StandardFile_Date.docx is our template. inside are two tables that will need some input that should be taken from the Application.
Some general Information has to put in above. This is using these Variables:
    - {{theme}} -> Name of the Planner
    - {{name}} -> Name of Me (should be changeable in some kind of settings menu)
    - {{playNumber}} -> Number of Players (should be adjustable in some kind of settings menu)
    - {{tools}} -> combination of Tools per exercice (declared later)
    - {{time}} -> Time configured at startup of Planner
    - {{requirements}} -> Information about the Players (should be adjustable in some kind of settings menu)
    - {{team}} -> Information about the Team (should be adjustable in some kind of settings menu)

The lower table has to be expanded accordingly depending on how many Segments our Planner has.
One Row always consists of these Elements although they can be empty:
    - {{slotX.time}} -> Timeslot of the Segment
    - {{slotX.goal}} -> General Information (Segment Name)
    - {{slotX.organisation}} -> Organisationform of the Exercice (Dropdown Menu in App "Spielform" or "Übungsform")
    - {{slotX.explanation}} -> Exlpanation of the Exercice done in this Segment (should be adjustable when creating new Segment)
    - {{slotX.tools}} -> List of Tools (should be adjustable when creating new Segment)

For this to work some new stuff has to be implemented in the Application like a Settings Menu and a more advanced Segment creation.

Create yourself a proper list of tasks and then do them one by one.