# MsMailMover

Move a date range of Microsoft 365 messages from one folder to another with
Microsoft Graph. Its direct dependencies are `msal` for Microsoft sign-in and
`requests` for Graph HTTP calls.

## Entra setup

1. Create an **App registration** in the tenant that owns the mailbox.
2. In **Authentication**, enable **Allow public client flows**.
3. Add delegated Microsoft Graph permission **Mail.ReadWrite.Shared**. Grant admin
   consent when required by the tenant.
4. Set the IDs in your shell:

   ```sh
   export MS_TENANT_ID="your-tenant-id"
   export MS_CLIENT_ID="your-application-client-id"
   ```

Edit the settings block in `msmailmover.py`: mailbox, source and
destination folder paths, dates, and date fields. `SOURCES` is a list of
`(source folder, date field)` pairs. `END_DATE` is inclusive.
For top-level standard folders, use Graph's language-independent names such as
`sentitems`, `inbox`, and `archive`; nested folders still use their display
names (for example, `archive/2025`).

The script begins in dry-run mode. Review the output, then set `DRY_RUN = False`
to actually move messages:

```sh
uv run --env-file .env msmailmover.py
```

It displays a device-login URL and code. Authenticate as an account authorized
to access the configured mailbox. No access token is stored.

Before moving anything, the script collects all matching message IDs from every
page and source folder. Its final report compares the selected and successfully
moved counts for each folder and in total.
