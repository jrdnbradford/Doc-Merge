# Doc Merge
Google Sheet-bound document merge application built with Google Apps Script. 

The files in this repo are meant to be bound to a Google Sheet container. Information about data used for merging that could be helpful for troubleshooting issues is logged in the [Google Account](https://script.google.com/home/executions) that runs the application. This default behavior can be changed by switching the *logTroubleShootingInfo* boolean variable in [globals.gs](server/globals.gs) to *false*.

## Recommended OAuth Scopes
```json
{
    "oauthScopes": [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents",
        "https://www.googleapis.com/auth/spreadsheets.currentonly",
        "https://www.googleapis.com/auth/script.container.ui"
    ]
}
```

## Authors
**Jordan Bradford** - GitHub: [jrdnbradford](https://github.com/jrdnbradford)

## License
All code in this project is licensed under the MIT license. See [LICENSE.txt](LICENSE.txt) for details.
