async function getGraphAccessToken() {
    const exchangePostData = {
        "__type": "TokenRequest:#Exchange",
        "Resource": "https://graph.microsoft.com/"
    };
    const canary = await cookieStore.get("X-OWA-CANARY");
    
    console.log("Requesting an access token for Graph...")
    const tokenRequest = await fetch(
        "https://outlook.office.com/owa/service.svc?action=GetAccessTokenforResource",
        {
            "headers": {
                "action": "GetAccessTokenforResource",
                "content-type": "application/json; charset=utf-8",
                "x-owa-canary": canary.value,
                "x-owa-urlpostdata": encodeURIComponent(JSON.stringify(exchangePostData)),
                "x-req-source": "Mail"
            },
            "method": "POST",
            "mode": "cors",
            "credentials": "include"
        }
    );
    const { AccessToken, AccessTokenExpiry } = await tokenRequest.json();
    console.log(
        `%cSuccessfully retrieved token %c(expires: ${AccessTokenExpiry})`,
        "color: green",
        "color: gray; font-style: italic"
    )
    return AccessToken;
}

async function* scrollDirectory(graphToken, skip=0) {
    const pageSize = 200;
    const queryParams = new URLSearchParams({
        $top: pageSize,
        $skip: skip,
        $orderBy: "displayName",
        $filter: "personType/subclass eq 'OrganizationUser'"
    });
    const graphRequest = await fetch(
        "https://graph.microsoft.com/v1.0/me/people?" + queryParams.toString(),
        {
            "headers": {
                "authorization": `Bearer ${graphToken}`,
                "content-type": "application/json",
                "X-PeopleQuery-QuerySources": "Mailbox,Directory"
            },
            "method": "GET",
            "credentials": "omit"
        }
    );
    const graphData = await graphRequest.json();
    const people = graphData.value;
    console.log(
        `%cFetched %c${people.length}%c people, skipped ${skip}`,
        "color: green",
        "font-weight: bold",
        "font-weight: normal"
    )
    yield *people;
    if (graphData["@odata.nextLink"]) {
        yield *scrollDirectory(graphToken, skip + people.length);
    }
}

const shouldSkip = (graphPerson) => {
    // Generally indicates not an individual
    return !graphPerson.surname || !graphPerson.givenName;
}

const flattenPerson = (graphPerson) => ({
    email: graphPerson.scoredEmailAddresses[0].address,
    name: graphPerson.givenName,
    department: graphPerson.department,
    jobTitle: graphPerson.jobTitle
});

const objectToCsvRow = (row) => Object.values(row).map(value => {
    return typeof value === 'string' ? JSON.stringify(value) : value
}).toString() + "\n";

async function main() {
    const fileHandle = await window.showSaveFilePicker({
        suggestedName: "organisation-directory.csv",
        types: [
            {
                description: "CSV (Comma-separated values)",
                accept: { "text/csv": [".csv"] },
            }
        ]
    });
    const writableStream = await fileHandle.createWritable();
    const graphToken = await getGraphAccessToken();
    let n = 0;

    for await (const graphPerson of scrollDirectory(graphToken)) {
        const person = flattenPerson(graphPerson);

        if (shouldSkip(graphPerson)) {
            console.log(
                `%cSkipping ${person.email}`,
                "color: gray; font-style: italic; font-size: 0.5rem"
            );
            continue;
        }

        if (n === 0) {
            await writableStream.write(Object.keys(person).toString() + "\n")
        }
        
        await writableStream.write(objectToCsvRow(person));
        n += 1;
    }
    await writableStream.close();
    console.log(
        `%cWrote ${n} people to ${fileHandle.name}`,
        "color: green; font-weight: bold")
}
main()