function base64DecToArr(base64String) {
  let encodedString = base64String.replace(/-/g, "+").replace(/_/g, "/");
  switch (encodedString.length % 4) {
    case 0:
      break;
    case 2:
      encodedString += "==";
      break;
    case 3:
      encodedString += "=";
      break;
    default:
      throw createBrowserAuthError(BrowserAuthErrorCodes.invalidBase64String);
  }
  const binString = atob(encodedString);
  return Uint8Array.from(binString, (m) => m.codePointAt(0) || 0);
}

async function getGraphAccessToken() {
  // It is easier to get the token from the MSAL cache than it is to ask for a new one
  const { value } = await cookieStore.get("msal.cache.encryption");
  const encryption = JSON.parse(decodeURIComponent(value));

  const msalKey = Object.keys(localStorage).find((k) =>
    k.startsWith("msal.token.keys")
  );
  const msalTokenKeys = JSON.parse(localStorage.getItem(msalKey));
  const accessTokenKey = msalTokenKeys.accessToken.find((k) =>
    k.includes("https://graph.microsoft.com/user.readbasic.all")
  );
  const encryptedToken = JSON.parse(localStorage.getItem(accessTokenKey));

  const data = base64DecToArr(encryptedToken.data);
  const nonce = base64DecToArr(encryptedToken.nonce);
  const clientId = msalKey.split(".")[3];
  const baseKey = await window.crypto.subtle.importKey(
    "raw",
    base64DecToArr(encryption.key),
    "HKDF",
    false,
    ["deriveKey"]
  );
  const derivedKey = await window.crypto.subtle.deriveKey(
    {
      name: "HKDF",
      salt: nonce,
      hash: "SHA-256",
      info: new TextEncoder().encode(clientId),
    },
    baseKey,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
  const decrypted = await window.crypto.subtle.decrypt(
    { name: "AES-GCM", iv: new Uint8Array(12) },
    derivedKey,
    data
  );

  return JSON.parse(new TextDecoder().decode(decrypted)).secret;
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
