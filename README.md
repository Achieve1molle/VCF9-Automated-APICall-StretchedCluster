# VCF 9 Stretched Cluster Automation

**Production Version:** v2.0.6  
**Script Name:** `VCF9-StretchCluster-Automation-v2.0.6.ps1`  
**Author:** Michael Molle  
**Runtime:** PowerShell 7+ / WPF  
**Primary Use Case:** Generate, validate, and execute VCF 9 Management Domain stretched-cluster JSON payloads.

## Overview

VCF 9 Stretched Cluster Automation is a WPF-based PowerShell tool for preparing and submitting the `clusterStretchSpec` workflow used to stretch a VCF 9 management cluster across AZ1 and AZ2.

The tool connects to:

- **SDDC Manager** for cluster, host, validation, and execution APIs.
- **vCenter** for reliable vDS and vmnic discovery when SDDC Manager network-query APIs are unavailable.
- **NSX Manager** for pre-validation of TEP IP pool overlaps and NSX uplink profile consistency.

Version **2.0.6** is the current production release. It includes the v2.0 production validation and payload-safety capabilities plus UI, workflow, status, witness-validation, and background-task improvements added after the initial v2.0 release.

## Key Capabilities

### SDDC Manager integration

- Authenticates to SDDC Manager using `/v1/tokens`.
- Lists VCF clusters.
- Resolves AZ2 host IDs from SDDC Manager inventory.
- Generates the final `clusterStretchSpec` payload.
- Generates a validation wrapper payload.
- Submits validation with `POST /v1/clusters/{clusterId}/validations`.
- Executes stretch with `PATCH /v1/clusters/{clusterId}`.

### vCenter integration

- Verifies vCenter connectivity using PowerCLI.
- Uses vCenter host proxy-switch data as the reliable fallback source for:
  - full vDS names,
  - vDS count,
  - host vmnic-to-vDS mapping.
- Verifies that the Witness FQDN is present in vCenter inventory before Generate, Validate, or Execute.

### NSX Manager integration

- Connects to NSX Manager using API credentials.
- Checks AZ2 TEP range against existing NSX IP pool ranges.
- Blocks Generate, Validate, and Execute when the requested TEP range overlaps an existing NSX range.
- Suggests the next available range inside the supplied CIDR when possible.
- Queries NSX uplink profiles.
- Validates requested active and standby uplink names against NSX-known uplink names.
- Warns when the generated uplink profile name already exists.
- Warns when the selected VLAN or teaming policy does not appear in extracted NSX uplink profile inventory.

## Version 2.0.6 Change Summary

### UI and visual updates

- Updated the UI to the common dark/black operational theme used by the disconnected software depot tool.
- Removed the large duplicate in-window title because the script title is already displayed in the window title bar.
- Added a bottom workflow status indicator showing `Ready`, `Running`, or `Failed`.
- Added muted status colors to reduce the overly bright/neon appearance.
- Preserved the timestamped run folder and live log pane behavior.

### Responsiveness and background tasks

- Added busy-state handling so buttons are disabled while a task is active.
- Added heartbeat/status updates so administrators can see that the script is running and not frozen.
- Moved prerequisite module installation actions into background jobs:
  - Install ImportExcel,
  - Install VMware.PowerCLI,
  - Install VCF.PowerCLI.
- Added automatic prerequisite recheck after a background module install completes.
- Added NuGet provider handling for prerequisite module installation.
- Moved Validate and Execute SDDC API submissions into background jobs to keep the WPF interface responsive during longer API calls.

### Witness validation improvements

Before Generate, Validate, or Execute, the tool now performs two Witness checks:

1. **DNS resolution check** using the Witness FQDN.
2. **vCenter inventory check** to verify the Witness is present as an added host.

These checks help prevent workflow failures caused by Witness FQDN typos or missing Witness inventory entries.

### Operator-message cleanup

- Generate success popup now shows: `Successfully Generated Specification`.
- Validate success popup now shows: `Validation Success`.
- Execute success popup now shows: `Task started. Monitor progress in VCF Operations Tasks.`
- Failure popups are simplified and direct administrators to the log instead of displaying long generated file paths.

### Input examples and defaults

- AZ1 name is prefilled with `AZ1`.
- AZ2 name is prefilled with `AZ2`.
- AZ2 TEP CIDR includes an example format, such as `10.52.56.0/24`.
- Witness CIDR includes an example format, such as `10.52.57.0/24`.

## Version 2.0 Base Change Summary

### Validation and payload safety

- Added client-side IPv4 validation.
- Added CIDR validation.
- Added gateway-in-CIDR validation.
- Added TEP range start/end validation.
- Added range-order validation.
- Added VLAN range validation.
- Added NSX IP pool overlap detection.
- Added suggested available TEP range calculation.
- Added NSX uplink profile validation.
- Added generated uplink profile name conflict warning.
- Removed storage architecture selection because OSA/ESA is not represented in the stretch JSON payload and is inferred by VCF from the selected cluster.

### Network discovery improvements

- Keeps vCenter as the source for vDS/vmnic mapping.
- Handles SDDC Manager network-query API unavailability cleanly.
- Logs fallback detection without alarming 400/404 endpoint noise.
- Preserves full vDS names and refuses unsafe short vDS labels.

### Operational behavior

- Generate, Validate, and Execute all run pre-checks before building or submitting payloads.
- NSX checks run automatically when NSX Manager is connected.
- If NSX Manager is not connected, the workflow logs a warning and continues with local and SDDC Manager validation.
- All outputs are written to a timestamped run folder.

## Generated Files

Each run creates a timestamped output folder such as:

```text
VCFStretch-Run-yyyyMMdd-HHmmss
```

Typical output files include:

```text
VCFStretch-yyyyMMdd-HHmmss.log
DetectedNetworkRaw_<clusterId>_<timestamp>.json
clusterStretchSpec_<clusterId>_<timestamp>.json
clusterUpdateSpec_validationWrapper_<clusterId>_<timestamp>.json
ValidationResponse_<clusterId>_<timestamp>.json
ExecuteResponse_<clusterId>_<timestamp>.json
```

## End-to-End Workflow

```mermaid
sequenceDiagram
    autonumber
    participant Runner as Automation Host<br/>(PowerShell 7 / WPF)
    participant SDDC as SDDC Manager<br/>(VCF API)
    participant VC as vCenter<br/>(PowerCLI)
    participant NSX as NSX Manager<br/>(Policy / Manager API)
    participant Disk as Run Folder<br/>(JSON + Logs)

    Runner->>Runner: Recheck prerequisites<br/>or install modules in background
    Runner->>SDDC: POST /v1/tokens<br/>SDDC credentials
    SDDC-->>Runner: Bearer access token
    Runner->>SDDC: GET /v1/clusters<br/>GET /v1/hosts
    SDDC-->>Runner: Cluster and host inventory
    Runner->>VC: Connect-VIServer<br/>Verify vCenter
    VC-->>Runner: vCenter session established
    Runner->>SDDC: Attempt cluster network query
    SDDC-->>Runner: Query unavailable / unsupported
    Runner->>VC: Read host proxy switches<br/>vDS + vmnic mapping
    VC-->>Runner: Full vDS names and vmnic mapping
    Runner->>NSX: Connect NSX Manager
    NSX-->>Runner: NSX API session available
    Runner->>Runner: Operator enters AZ2 TEP pool,<br/>AZ names, hosts, witness values
    Runner->>Runner: Local validation<br/>IPv4, CIDR, gateway, range, VLAN
    Runner->>Runner: Witness DNS resolution check
    Runner->>VC: Verify Witness host in vCenter inventory
    VC-->>Runner: Witness inventory result
    Runner->>NSX: Query existing IP pools
    NSX-->>Runner: Existing pool ranges
    Runner->>Runner: Block if TEP range overlaps<br/>Suggest next available range
    Runner->>NSX: Query uplink profiles
    NSX-->>Runner: Uplink profile inventory
    Runner->>Runner: Validate active/standby uplinks<br/>Warn on profile-name conflict
    Runner->>Disk: Write clusterStretchSpec JSON
    Runner->>Disk: Write validation wrapper JSON
    Runner->>SDDC: POST /v1/clusters/{clusterId}/validations<br/>background job
    SDDC-->>Runner: Validation task / response
    Runner->>Disk: Save ValidationResponse JSON
    alt Validation succeeded and operator confirms Execute
        Runner->>SDDC: PATCH /v1/clusters/{clusterId}<br/>background job
        SDDC-->>Runner: Execute response / task
        Runner->>Disk: Save ExecuteResponse JSON
    else Validation failed or pre-check failed
        Runner->>Runner: Correct inputs and re-run validation
    end
```

## Prerequisites

### Workstation

- Windows workstation or management VM.
- PowerShell 7+.
- .NET/WPF support.
- Network reachability to:
  - SDDC Manager,
  - vCenter,
  - NSX Manager,
  - Witness appliance,
  - AZ2 ESXi hosts.

### PowerShell modules

Recommended modules:

```text
VMware.PowerCLI
VCF.PowerCLI
ImportExcel
```

The UI includes buttons to install or re-check common prerequisites. In v2.0.6, module installation runs in the background so the WPF interface remains responsive.

### VCF / SDDC prerequisites

- Target cluster selected in SDDC Manager.
- AZ2 hosts commissioned and visible in SDDC Manager.
- Witness appliance deployed and reachable.
- Witness FQDN resolves in DNS.
- Witness host is present in vCenter inventory.
- Witness FQDN, vSAN IP, and vSAN CIDR known.
- AZ2 NSX TEP VLAN, CIDR, gateway, and range planned.
- NSX TEP range confirmed not to overlap existing pools.

## How to Run

Launch from PowerShell 7:

```powershell
pwsh.exe -ExecutionPolicy Bypass -File .\VCF9-StretchCluster-Automation-v2.0.6.ps1
```

Recommended workflow:

1. Click **Recheck** in Prerequisites.
2. If modules are missing, click the relevant install button and wait for the status to return to **Ready**.
3. Enter SDDC Manager FQDN and credentials.
4. Click **Connect SDDC**.
5. Select the target cluster.
6. Enter vCenter FQDN and credentials.
7. Click **Verify vCenter**.
8. Enter NSX Manager FQDN and credentials.
9. Click **Connect NSX**.
10. Click **Detect Network**.
11. Fill AZ2 TEP pool values.
12. Fill AZ1/AZ2 names, AZ2 hosts, and Witness details.
13. Expand Advanced only if you need to override generated uplink values.
14. Click **Generate JSON**.
15. Click **Validate**.
16. Review validation output and the generated response file.
17. Click **Execute** only during the approved change window.
18. Monitor progress in **VCF Operations Tasks**.

## Main UI Sections

### Prerequisites

Displays PowerShell and module status. Module installation actions run in background jobs in v2.0.6.

### Connections

Collects connection details for:

- SDDC Manager,
- vCenter,
- NSX Manager.

Passwords are not saved in configuration files.

### Cluster and Detection

Contains:

- Cluster selector.
- Detect Network button.
- Stacked detection summary:
  - vDS Count,
  - vDS Names,
  - Uplinks.

### AZ2 NSX TEP Pool

Collects:

- pool name,
- CIDR,
- gateway,
- range start,
- range end,
- transport VLAN.

### Advanced Generated Names / Uplink Profile

Contains generated and override-capable fields:

- generated AZ2 network profile,
- uplink profile,
- teaming policy,
- active uplinks,
- standby uplinks,
- vDS-to-NSX uplink mapping.

Most users should not need to modify these unless NSX validation reports a mismatch.

### Stretch Inputs

Collects:

- AZ1 name,
- AZ2 name,
- AZ2 host FQDNs,
- Witness FQDN,
- Witness vSAN IP,
- Witness vSAN CIDR,
- license and Edge multi-AZ options.

## Validation Behavior

Before JSON generation, validation, or execution, the script performs:

- Required field checks.
- IPv4/CIDR validation.
- Gateway-in-CIDR validation.
- TEP range-in-CIDR validation.
- TEP start/end order validation.
- VLAN validation.
- Witness FQDN DNS resolution check.
- Witness host presence check in vCenter inventory.
- NSX IP pool overlap validation.
- NSX uplink profile validation.
- Host ID resolution through SDDC Manager.
- vDS consistency checks based on detected vCenter data.

If NSX Manager is connected, NSX checks are enforced. If NSX Manager is not connected, the script logs a warning and continues with local and SDDC Manager validation.

## Troubleshooting

### Module install appears slow

In v2.0.6, module installs run in the background. The UI should remain responsive and the status should show `Running`. If installation still takes a long time, check network, proxy, DNS, TLS, and PowerShell Gallery reachability.

### Witness FQDN validation fails

Confirm that the Witness FQDN is spelled correctly, resolves in DNS, and is present in vCenter inventory as a host. Correct the Witness value and rerun Generate or Validate.

### NSX IP overlap failure

Example:

```text
Requested TEP range 10.52.56.200-10.52.56.210 overlaps existing NSX range(s):
10.52.56.194-10.52.56.201, 10.52.56.202-10.52.56.205.
Suggested available range: 10.52.56.206-10.52.56.216.
```

Correct the range and re-run Generate or Validate.

### NSX uplink mismatch

If validation reports that `uplink1` or `uplink2` is not found in NSX-known uplink names, expand Advanced and adjust:

```text
Active Uplinks
vDS -> NSX Map
```

Common alternatives include:

```text
uplink-1,uplink-2
```

or:

```text
uplink1,uplink2
```

Use the naming convention reported by NSX validation.

### SDDC Manager network query unavailable

This is expected in some environments. The script uses vCenter PowerCLI fallback for vDS and vmnic discovery.

### UI log is empty

The script includes dispatcher-safe live log updates. If the UI log does not update, check the run folder log file.

## Security Notes

- SDDC Manager bearer tokens are held in memory only.
- Passwords are not saved in configuration files.
- Generated JSON and API responses may contain environment-specific infrastructure data.
- Store run folders securely.
- Review payloads before Execute.
- Execute should be used only inside an approved change window.

## License

Internal use. Provide attribution if reused or modified.
