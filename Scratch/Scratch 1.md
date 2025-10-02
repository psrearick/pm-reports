The "Docs/Architecture.md," "Docs/System Specification.md," and "Docs/Process.md," give you an idea of what my project is. I am working on a Google Apps sheet project and working on making a template that loops through rows sort and group the rows. What I have is a table where the leftmost column is "unit" and it will have values like "E-5" and "A-1" and there is also a "date" column with dates.

I start the loop with this line in a cell in column A on an empty row: `[[ROW transactions group=unit sort=date]]`

Then, I define the row fields, and I end the table with: `[[ENDROW]]`

I want all the rows of a specific unit to be together. I also want the list sorted by unit. Then, I want it to be sorted by date within each grouped unit entries.

Within each group, the grouped column should only have a value for the first item in the group. So every "E-5" unit will be listed together, but only the first one will have a value in the "unit" column.

The logic for this is in "Apps Script/TemplateEngine.gs" but it is not working.

There will by a row with a specific unit, like "E-5" and the following rows, which are also E-5 entries will not list the unit, which is correct. Then there will be a different unit below. Then later on there will be more entries for E-5. This is incorrect because all E-5 entries should be together.

It is correctly sorting by date if I only sort by date without grouping. (`[[ROW transactions sort=date]]`)

It is correctly sorting by unit if I only sort by date without grouping.
(`[[ROW transactionssort=unit]]`)

 It groups by unit correctly, but if it is not also sorted by unit, there may be multiple groups for the same unit. (`[[ROW transactions group=unit]]`). If it is only sorted by unit, then it will not be sorted by date.

Additionally, grouping by unit and sorting by unit (`[[ROW transactions group=unit sort=unit]]`) does not do the grouping.

Based on this information, look through this projects code and fix the issue.
