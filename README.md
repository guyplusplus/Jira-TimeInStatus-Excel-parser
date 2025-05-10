# Jira-TimeInStatus-Excel-parser
Excel macro that parses for each issue the field 'Time in Status' and creates columns '[TiS] status count' and '[TiS] status duration'.

Steps are:
1. Export issues as Excel csv
2. Open the CSV file with Excel
3. Add this macro (ALT-F11)
4. Run this macro (ALT-F8)

The macro adds new columns. The output will look like this:

| Custom field ([CHART] Time in Status) | [TiS] OPEN - Count | [TiS] OPEN - Duration | [TiS] Req in progress - Count | [TiS] Req in progress - Duration | ... |
| --- | --- | --- | --- | --- | --- |
| 1\_\*:\*\_1\_\*:\*\_15888\_\*\|\*\_3\_\*:\*\_1\_\*:\*\_0\_\*\|\*\_10000\_\*:\*\_1\_\*:\*\_34198 | 1 | 0 days 00:00:16 | | | ... |
| 1\_\*:\*\_1\_\*:\*\_580626\_\*\|\*\_10390\_\*:\*\_1\_\*:\*\_13406\_\*\|\*\_3\_\*:\*\_1\_\*:\*\_2147\_\*\|\*\_10000\_\*:\*\_1\_\*:\*\_1767\_\*\|\*\_10387\_\*:\*\_3\_\*:\*\_2198\_\*\|\*\_10277\_\*:\*\_1\_\*:\*\_1666 | 1 | 0 days 00:09:41 | 3 | 0 days 00:00:02 | ... |

To get the list of statuses in a give Atlassian site, you can use this command below and update the macro and add necessary keys (staus ID, status name).

```shell
.\curl.exe -H "Accept: application/json" --user john@gmail.com:123mytoken789 https://mysite.atlassian.net/rest/api/3/status
```

```json
[
    {
        "description": "",
        "iconUrl": "https://mysite.atlassian.net/images/icons/statuses/open.png",
        "id": "1",
        "name": "OPEN",
        "self": "https://mysite.atlassian.net/rest/api/3/status/1",
        "statusCategory": {
            "colorName": "blue-gray",
            "id": 2,
            "key": "new",
            "name": "To Do",
            "self": "https://mysite.atlassian.net/rest/api/3/statuscategory/2"
        },
        "untranslatedName": "OPEN"
    },
    {
        "description": "This issue is being actively worked on at the moment by the assignee.",
        "iconUrl": "https://mysite.atlassian.net/images/icons/statuses/inprogress.png",
        "id": "3",
        "name": "In Progress",
        "self": "https://mysite.atlassian.net/rest/api/3/status/3",
        "statusCategory": {
            "colorName": "yellow",
            "id": 4,
            "key": "indeterminate",
            "name": "In Progress",
            "self": "https://mysite.atlassian.net/rest/api/3/statuscategory/4"
        },
        "untranslatedName": "In Progress"
    },
    {
        "description": "This issue was once resolved, but the resolution was deemed incorrect. From here issues are either marked assigned or resolved.",
        "iconUrl": "https://mysite.atlassian.net/images/icons/statuses/reopened.png",
        "id": "4",
        "name": "Reopened",
        "self": "https://mysite.atlassian.net/rest/api/3/status/4",
        "statusCategory": {
            "colorName": "blue-gray",
            "id": 2,
            "key": "new",
            "name": "To Do",
            "self": "https://mysite.atlassian.net/rest/api/3/statuscategory/2"
        },
        "untranslatedName": "Reopened"
    },
]
```