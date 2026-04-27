# How to Get an Azure DevOps Personal Access Token (PAT)

A PAT is a secure password that lets the FOXL Deploy tool authenticate to Azure DevOps
on your behalf. It is only stored in memory — never saved to disk.

---

## Step-by-step

1. **Open the Azure DevOps token page**

   https://dev.azure.com/Ninety-One/_usersSettings/tokens

   (Sign in with your Ninety One Microsoft account if prompted.)

2. **Click "New Token"** (top-right of the page).

3. **Fill in the token details:**

   | Field        | Value                        |
   |--------------|------------------------------|
   | Name         | `FOXL Deploy` (or anything)  |
   | Organization | `Ninety-One`                 |
   | Expiration   | Up to 1 year — your choice   |
   | Scopes       | **Custom defined**           |

4. **Set the scope** — expand "Build" and tick **Read** only.
   No other scopes are needed.

5. **Click "Create".**

6. **Copy the token immediately.** Azure DevOps shows it only once.
   Paste it into the PAT field in the Deploy screen.

---

## Notes

- The token expires on the date you set. If "Load Builds" returns a 401 error,
  your token has likely expired — generate a new one using the steps above.
- Never share your PAT or commit it to source control.
- To revoke a token at any time, return to the tokens page and click "Revoke".

---

## Troubleshooting

| Error                  | Likely cause                          |
|------------------------|---------------------------------------|
| HTTP 401               | Wrong or expired PAT                  |
| HTTP 403               | PAT lacks Build (Read) scope          |
| Connection error       | No network / VPN not connected        |
| "No artifact URL"      | Build did not produce a Binaries5 artifact (e.g. failed build) |
