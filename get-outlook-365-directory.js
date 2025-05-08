async function getGraphAccessToken() {
  const exchangePostData = {
    __type: "TokenRequest:#Exchange",
    Resource: "https://graph.microsoft.com/",
  };

  console.log("Requesting an access token for Graph...");
  const tokenRequest = await fetch(
    "https://outlook.office.com/owa/service.svc?action=GetAccessTokenforResource",
    {
      headers: {
        action: "GetAccessTokenforResource",
        "content-type": "application/json; charset=utf-8",
        "x-owa-urlpostdata": encodeURIComponent(
          JSON.stringify(exchangePostData)
        ),
        "x-req-source": "Mail",
      },
      method: "POST",
      mode: "cors",
      credentials: "include",
    }
  );
  const { AccessToken, AccessTokenExpiry } = await tokenRequest.json();
  console.log(
    `%cSuccessfully retrieved token %c(expires: ${AccessTokenExpiry})`,
    "color: green",
    "color: gray; font-style: italic"
  );
  return AccessToken;
}

async function* scrollDirectory(graphToken) {
  const pageSize = 100;
  const headers = {
    authorization: `Bearer ${graphToken}`,
    "content-type": "application/json",
    ConsistencyLevel: "eventual",
  };
  const initialParams = new URLSearchParams({
    $count: true,
    $top: pageSize,
    $orderby: "displayName",
    $select: "displayName,jobTitle,department,mail",
    $filter: "jobTitle ne null and accountEnabled eq true",
  });

  async function* getPage(url) {
    const res = await fetch(url, {
      headers,
      method: "GET",
      credentials: "omit",
    });
    const data = await res.json();
    console.log(
      `%cFetched %c${data.value.length}%c people`,
      "color: green",
      "font-weight: bold",
      "font-weight: normal"
    );
    yield* data.value;

    const nextPage = data["@odata.nextLink"];
    if (nextPage) {
      yield* getPage(nextPage);
    }
  }

  yield* getPage(
    `https://graph.microsoft.com/v1.0/users?${initialParams.toString()}`
  );
}

const objectToCsvRow = (row) =>
  Object.values(row)
    .map((value) => {
      return typeof value === "string" ? JSON.stringify(value) : value;
    })
    .toString() + "\n";

async function main() {
  if (window.location.host !== "outlook.office.com") {
    alert("This needs to be run while you're on outlook.office.com");
    return;
  }

  const fileHandle = await window.showSaveFilePicker({
    suggestedName: "organisation-directory.csv",
    types: [
      {
        description: "CSV (Comma-separated values)",
        accept: { "text/csv": [".csv"] },
      },
    ],
  });
  const writableStream = await fileHandle.createWritable();
  const graphToken = await getGraphAccessToken();
  let n = 0;

  for await (const user of scrollDirectory(graphToken)) {
    if (n === 0) {
      await writableStream.write(Object.keys(user).toString() + "\n");
    }

    await writableStream.write(objectToCsvRow(user));
    n += 1;
  }
  await writableStream.close();
  console.log(
    `%cWrote ${n} users to ${fileHandle.name}`,
    "color: green; font-weight: bold"
  );
  alert(`Wrote ${n} users to ${fileHandle.name}`);
}
main();
