Please readin a json file (players.json) as input in #index.js, it's formated as: 

```JSON
[
    {
        "omnipongId": "111538",
        "omnipongNote": "Registered on 05/11/2018",
        "usattId": "223250",
        "usattNotes": "Expires on 09/13/2025",
        "rating": "1122",
        "ratingNote": "Last Updated from 09/24/2023",
        "lastName": "Abazia",
        "firstName": "Jared",
        "address": "553 Winding River Court",
        "city": "Brick",
        "state": "",
        "email": "jabazia@gmail.com",
        "phone": "848-666-3717",
        "club": "None",
        "dob": "10/08/2002",
        "gender": "",
        "fileId": {}
    }
]
```

use it's content to pupulate a new xlsx file into tournaments\YYYY-MM-DD\justgo-import.xlsx, using the template as #JustGo_Member_Import.xlsx, which has header columns: 
Firstname*	Lastname*	EmailAddress*	DOB*	Username*	Gender	Title	Address1	Address2	Town	County	PostCode	Country	Mobile Telephone	Home Telephone	Emergency Contact First Name	Emergency Contact Surname	Emergency Contact Relationship	Emergency Contact Number	Emergency Contact Email Address	Parent FirstName	Parent Surname	Parent EmailAddress

where the one with * must have.