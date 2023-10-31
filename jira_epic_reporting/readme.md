
# Jira Reporting

## Command Line For Project Report

### ReviewMarketing Label with Console Output and Links in review_marketing.txt

```
python3 project_report.py --label ReviewMarketing --console --file review_marketing.txt
```

reviewmarketing.json will be a control file that contains
- Sprint Start Date
- Sprint End Date

```
{
    "start_sprint": "155",
    "end_sprint": "160"
}
```

### ReviewMarketing Label with Console Output and Links in review_marketing.txt and create ai review tab

```
python3 project_report.py --label ReviewMarketing --console --file review_marketing.txt --ai
```

### Trail Project Label 

```
python3 project_report.py --label ReportTrial --console --file review_marketing.txt
```

## Command Line for Jira Information

Will get all boards, sprints, users, projects in Jira

```
python3 project_report.py --info
```

### Example .env file

```
JIRA_API_KEY=amhvbG1lc0ByZXZsb2NhbC5jb206Y1ZKZm1aSWhRMm94YzltaXo3aTkwMEI0
CLAUDE_API_KEY=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
JIRA_REV_LOCATION=https://revlocaldev.atlassian.net
JIRA_SEARCH=rest/api/2/search?jql=
JIRA_BOARD=rest/agile/1.0/board/
JIRA_ISSUE=rest/api/2/issue/
PATH_LOCATION=/Users/jholmes/Library/CloudStorage/OneDrive-RevLocal/reviewmarketing/
JIRA_ISSUE_LINK=https://revlocaldev.atlassian.net/browse/
```

[One Drive Output Location](https://revlocalinc-my.sharepoint.com/:f:/g/personal/jholmes_revlocal_com/EiR1Aui9R9ZEirrVwyOyLeIBHfm2fngUvXbFNcD-nczL3w?e=o8o1le)

## PDF File

[Real Python PDF Files](https://realpython.com/creating-modifying-pdf/)

### Claude

[Claude Prompt Chaining](https://docs.anthropic.com/claude/docs/prompt-chaining)

### Install Items for reporting

### Excel
openpyxl

### Console Color Ouput
colorama

### Environment Variables
dotenv

### Object Typing
typing

### Claude Usage
anthropic

### PDF Creation

### Chart Creation

[Scatter Chart](https://stackoverflow.com/questions/63696835/python-openpyxl-scatter-plots-with-secondary-y-axis)

### Chart Background

[Chart Background](https://www.youtube.com/watch?v=osFd-8B146w)

### Fill Gradient

[fills](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fills.html)

### Cell Formatting Examples
[Cell Formatting](https://www.blog.pythonlibrary.org/2021/08/11/styling-excel-cells-with-openpyxl-and-python/)

### HTML Color Code Selector

[Color Code Selector](https://www.rapidtables.com)
