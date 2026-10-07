# EntraDeviceCleanup

## Introduction

[Invoke-EntraDeviceCleanup.ps1](./Invoke-EntraDeviceCleanup.ps1) is an Azure Automation PowerShell runbook that reports on, disables, and deletes stale Microsoft Entra ID device registrations. It uses Microsoft Graph to identify inactive devices and checks Microsoft Intune enrollment before deleting them.

The script defaults to **read-only mode**: it reports proposed actions without changing devices. Optionally, it sends a summary Adaptive Card to a Microsoft Teams Workflows webhook.

> Disabling devices can prevent access to resources, and deleting device registrations is destructive. Review the read-only results and agree on retention thresholds before enabling production mode. The script is provided "AS IS" with no warranties.

## How cleanup works

The script retrieves all Entra ID devices and Intune-managed devices, then classifies them using UTC dates:

1. Use the device's `approximateLastSignInDateTime` as its last activity date. If this is missing, use `registrationDateTime`. Devices without either date are ignored.
2. Ignore devices newer than the disable threshold.
3. Exclude hybrid joined devices (`trustType = ServerAd`) when `CleanUPExcludeHybrid` is enabled.
4. Always exclude devices whose `physicalIds` contain an Autopilot `[ZTDID]` marker.
5. Mark enabled devices at or beyond the disable threshold for disabling.
6. Mark already-disabled devices at or beyond the delete threshold for deletion, unless their Entra `deviceId` matches an Intune managed device's `azureADDeviceId`.

**Important behavior:**

- Thresholds are inclusive: a device whose last activity is exactly at a cutoff is eligible.
- The delete threshold measures inactivity, **not time since the device was disabled**. An enabled device is only disabled during that run; a later run may delete it if it already meets the delete threshold.
- Intune-managed devices are protected from **deletion only**. Stale enabled Intune-managed devices can still be disabled.
- Autopilot protection is based on the marker on the Entra device, not a separate Autopilot inventory query.
- If the number of delete candidates exceeds `CleanUPMaxDeletes`, **all deletions are skipped** for that run. Disabling still proceeds in production mode.
- Per-device failures are recorded while remaining actions continue. The job throws an error at the end if any device action failed.

## Prerequisites

- An Azure subscription and an Azure Automation account.
- A PowerShell **7.4** runtime environment for the runbook.
- These Microsoft Graph PowerShell modules installed in that runtime environment:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Identity.DirectoryManagement`
  - `Microsoft.Graph.DeviceManagement`
- A Microsoft Entra ID app registration with a certificate for unattended, app-only authentication.
- Permission to configure the Automation account, register an application, grant tenant-wide admin consent, and assign the required Entra role.
- A certificate with its private key available to the runbook process and its public certificate uploaded to the app registration.
- Optionally, a Teams Workflow that accepts an HTTP webhook request and posts the supplied Adaptive Card to a channel or chat.

No interactive user sign-in or client secret is used by the certificate authentication path.

## Configure the Entra ID app registration

1. In the Microsoft Entra admin center, open **App registrations > New registration**.
2. Create a single-tenant application, for example `EntraDeviceCleanup`. A redirect URI is not required.
3. Record its **Application (client) ID** and **Directory (tenant) ID**.
4. Under **API permissions > Add a permission > Microsoft Graph > Application permissions**, add:

   | Application permission | Purpose |
   | --- | --- |
   | `Device.ReadWrite.All` | Read Entra devices, disable them, and delete registrations. |
   | `DeviceManagementManagedDevices.Read.All` | Read Intune-managed devices to exclude them from deletion. |

5. Grant administrator consent for your tenant. These are **application**, not delegated, permissions.
6. Under **Roles & admins > Cloud Device Administrator > Add assignments**, assign the application's **service principal** (the enterprise application) the Cloud Device Administrator role. The script expects this role for app-only device disable/delete operations. Assigning an Azure subscription/resource-group RBAC role is not a substitute.
7. Under the app registration's **Certificates & secrets > Certificates**, upload the public certificate (`.cer`). Never upload the private-key `.pfx` here.

Use the same application and tenant IDs in the Automation variables below.

## Configure certificate authentication in Azure Automation

### Create or obtain a certificate

Use a certificate approved by your organization. For an initial setup, the following example creates a self-signed certificate on a Windows administration workstation and exports a public certificate and a password-protected private-key certificate:

```powershell
$certificate = New-SelfSignedCertificate `
    -Subject 'CN=EntraDeviceCleanup' `
    -CertStoreLocation 'Cert:\CurrentUser\My' `
    -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider' `
    -KeySpec Signature `
    -KeyExportPolicy Exportable `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(1)

$password = Read-Host 'Enter a password for the PFX export' -AsSecureString
Export-Certificate -Cert $certificate -FilePath '.\EntraDeviceCleanup.cer'
Export-PfxCertificate -Cert $certificate -FilePath '.\EntraDeviceCleanup.pfx' -Password $password
$certificate.Thumbprint
```

Upload the `.cer` to the app registration. Protect the `.pfx` and its password as credentials, and plan certificate renewal before expiry.

### Make the private key available to the runbook

1. In the Automation account, open **Shared resources > Certificates > Add a certificate**.
2. Upload the `.pfx` with its password, using an asset name such as `EntraDeviceCleanupCert`. Use an exportable asset if your deployment needs to transfer the certificate into the process certificate store.
3. Set `CleanUPCert` to the certificate's **thumbprint**, not the Automation certificate asset name.

**The repository script uses `Connect-MgGraph -CertificateThumbprint`; it does not call `Get-AutomationCertificate`.** Uploading an Automation certificate asset alone is not a replacement for making that certificate and its private key discoverable in the job's certificate store.

For an Automation certificate asset, add the following bootstrap to the deployed runbook **before the repository script**. It retrieves the asset and places it in the current user's personal certificate store for thumbprint lookup:

```powershell
$automationCertificate = Get-AutomationCertificate -Name 'EntraDeviceCleanupCert' -ErrorAction Stop
if (-not $automationCertificate -or -not $automationCertificate.HasPrivateKey) {
    throw 'EntraDeviceCleanupCert must contain a private key.'
}

$configuredThumbprint = Get-AutomationVariable -Name 'CleanUPCert' -ErrorAction Stop
if ($automationCertificate.Thumbprint -ne $configuredThumbprint) {
    throw 'CleanUPCert does not match the EntraDeviceCleanupCert asset thumbprint.'
}

$store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
try {
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $store.Add($automationCertificate)
}
finally {
    $store.Close()
}
```

Alternatively, on a Windows Hybrid Runbook Worker, provision the certificate and private key in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My` for the worker's execution identity and ensure that identity can access the private key. A certificate installed only on your own workstation is not available to the runbook.

When renewing the certificate, update the app registration's public certificate, the Automation asset or worker certificate, and `CleanUPCert` together. Test authentication before retiring the old certificate.

## Automation variables

In the Automation account, open **Shared resources > Variables > Add a variable**. Create the following variables using the exact names and types shown. Defaults apply when a variable is absent or empty.

| Name | Type | Default | Purpose / example |
| --- | --- | --- | --- |
| `CleanUPAPPID` | String | None | Application (client) ID of the app registration. Required for certificate authentication. |
| `CleanUPTenantID` | String | None | Directory (tenant) ID. Required for certificate authentication. |
| `CleanUPCert` | String | None | Certificate thumbprint, without spaces. Required for certificate authentication; this is not the asset name or a file path. |
| `CleanUPReadOnly` | Boolean | `true` | `true` reports proposed actions only; `false` enables device disable/delete operations. |
| `CleanUPDisableDays` | Integer | `60` | Inactivity threshold for disabling enabled devices. |
| `CleanUPDeleteDays` | Integer | `90` | Inactivity threshold for deleting already-disabled devices. Must be greater than `CleanUPDisableDays`. |
| `CleanUPMaxDeletes` | Integer | `100` | Maximum permitted delete-candidate count per run. If exceeded, no candidates are deleted. |
| `CleanUPExcludeHybrid` | Boolean | `true` | Exclude hybrid joined (`ServerAd`) devices. Set to `false` only if you intentionally want them included. |
| `CleanUPTeamsURI` | String | None | Optional Teams Workflows webhook URL. Leave absent or empty to disable Teams notifications. Store as an **encrypted** variable. |

Recommended initial configuration:

- Populate all three certificate authentication variables.
- Keep `CleanUPReadOnly = true`.
- Set `CleanUPDisableDays = 60`, `CleanUPDeleteDays = 90`, `CleanUPMaxDeletes = 100`, and `CleanUPExcludeHybrid = true`, adjusting thresholds to your organization's policy.
- Use actual Boolean variables rather than strings for the Boolean settings.
- Use positive integers for day thresholds. Legacy negative day values are accepted and converted to their absolute values. The runbook stops if the delete threshold is not greater than the disable threshold.
- Use a nonnegative delete cap. A cap of `0` prevents deletion whenever there are delete candidates.

### Authentication selection and managed identity alternative

Certificate authentication is used **only when `CleanUPAPPID`, `CleanUPTenantID`, and `CleanUPCert` are all populated**. If any one is missing or empty, the script attempts `Connect-MgGraph -Identity` instead; it does not report an incomplete certificate configuration.

To intentionally use managed identity, leave all three authentication variables absent or empty, enable the Automation account's system-assigned managed identity, and grant its service principal the same Microsoft Graph application permissions and Cloud Device Administrator role. Graph application permissions must be assigned to that service principal; Azure RBAC assignments alone do not grant Graph access. Omit the certificate bootstrap when using managed identity.

## Deploy, test, and schedule the runbook

1. Create/select a PowerShell 7.4 runtime environment in Azure Automation and add the three Graph modules listed above. Wait for installation to complete.
2. Create a PowerShell runbook named `Invoke-EntraDeviceCleanup`, associate it with that environment, and import/paste the [script](./Invoke-EntraDeviceCleanup.ps1).
3. If using an Automation certificate asset, prepend the certificate bootstrap shown above to the deployed runbook.
4. Configure the certificate, app registration, permissions, role, and Automation variables.
5. Keep `CleanUPReadOnly = true` and run the **Test pane**. Verify the authentication method in the job output, inspect disable/delete candidates, check exclusions, and confirm any Teams notification.
6. Publish the runbook and run a read-only job to verify the published deployment.
7. After approving the results, set `CleanUPReadOnly = false`. This changes behavior on the next run without requiring script edits.
8. Create and link an Automation schedule, for example a daily run. Review job output and failures regularly, and monitor certificate expiry.

The script reads configuration from Automation variables; no runbook parameters are required. Runbook test and production jobs use the same variables, so changing `CleanUPReadOnly` also affects later test runs.

## Reporting and Teams notifications

Job output includes the device counts and tables for disable candidates, delete candidates, skipped Autopilot devices, skipped Intune deletion candidates, skipped hybrid joined devices, and failed actions. Exclusion counts cover devices considered by the stale-device classification, not the entire tenant inventory.

To enable Teams reporting, create a Teams Workflow using a webhook trigger that accepts the script's unauthenticated HTTP POST, and configure it to post the supplied Adaptive Card. Put its generated URL in the encrypted `CleanUPTeamsURI` variable. A trigger requiring a signed-in user or bearer token will not work with the script's webhook call. Treat the URL as a secret and restrict access to the workflow and destination.

The card includes the execution mode, thresholds, candidate counts, exclusions, failures, and up to 40 entries for each action list. Production action counts reflect **candidates**, not verified successful operations; check failed actions and job output before interpreting them as success totals. Failed webhook delivery generates warnings but does not by itself fail the cleanup job.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| Unexpected managed identity connection | Verify all three certificate authentication variables are populated. |
| Certificate not found or authentication fails | Verify the thumbprint, certificate store availability to the job, private-key access, expiry, tenant/client IDs, and matching public certificate on the app registration. |
| Graph access denied / `403` | Verify application permissions, tenant admin consent, and the Cloud Device Administrator assignment to the correct service principal. |
| Invalid retention thresholds | Ensure the absolute value of `CleanUPDeleteDays` is greater than that of `CleanUPDisableDays`. |
| Deletion cap warning | Review the full candidate list before changing the cap. The script skips all deletions, rather than deleting only the first candidates up to the cap. |
| Teams notification warning | Verify the webhook URL, workflow trigger authentication, workflow status, and Adaptive Card handling. |

## References

- [Azure Automation runtime environments](https://learn.microsoft.com/en-us/azure/automation/runtime-environment-overview)
- [Azure Automation certificates](https://learn.microsoft.com/en-us/azure/automation/shared-resources/certificates)
- [Azure Automation variables](https://learn.microsoft.com/en-us/azure/automation/shared-resources/variables)
- [Microsoft Graph PowerShell authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands)
