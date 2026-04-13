# VCF 9 Stretched Cluster Automation — Single‑Page Wiki

**Version:** v0.10.33  
**Author:** Michael Molle  

This utility is a WPF-based PowerShell 7 tool that helps you **generate, validate, and execute** the JSON payloads required to stretch a VCF 9 management cluster across AZ1 and AZ2. It reads from **SDDC Manager (source of truth)**, optionally connects to **vCenter**, and can also **populate an Excel data-collection workbook**.

---

## What the tool does
- Acquires an SDDC Manager **token** (`POST /v1/tokens`) and uses it for subsequent API calls.
- Lists clusters and hosts; lets you pick **UNASSIGNED/USEABLE** hosts for AZ2.
- **Collects** `VCENTER_NSXT_NETWORK_CONFIG` for the selected cluster and **auto‑fills** many NSX/vDS fields.
- Builds a complete `clusterStretchSpec` and writes:
  - `clusterStretchSpec_<clusterId>_<timestamp>.json`
  - `clusterUpdateSpec_validationWrapper_<clusterId>_<timestamp>.json` (wrapper for validation)  
  - Validation/Execution responses to timestamped JSON files  
  - Optional: fills your Excel workbook via **ImportExcel**
- **Validates** your spec (`POST /v1/clusters/{id}/validations`) and, if desired, **executes** the stretch (`PATCH /v1/clusters/{id}`).

---

## End‑to‑End Flow
mermaid


flowchart TD
  A[Start UI (PowerShell 7, WPF)] --> B[SDDC Manager: Connect\nPOST /v1/tokens]
  B --> C[Load Clusters / Hosts]
  C --> D{Optional: vCenter Connect\n(VCF.PowerCLI or VMware.PowerCLI)}
  D --> C
  C --> E[Collect VCENTER_NSXT_NETWORK_CONFIG]
  E --> F[Auto‑populate TEP pool, uplink profile,\ntransport VLAN, uplinks, vDS and mappings]
  F --> G[Select/Confirm AZ2 hosts\n(UNASSIGNED/USEABLE)]
  G --> H[Fill Witness (FQDN, vSAN IP/CIDR)]
  H --> I[Generate JSON payloads]
  I --> J[Validate: POST /v1/clusters/{id}/validations]
  J --> K{Validation OK?}
  K -- No --> F
  K -- Yes --> L[Execute: PATCH /v1/clusters/{id}]
  L --> M[Save responses & logs\nOpen Reports Folder]


---

## Connections & Tokens
### SDDC Manager (required)
- **Inputs:** FQDN, username (e.g., `administrator@vsphere.local`), password  
- **Auth:** `POST https://<sddc-manager>/v1/tokens` returns `accessToken` and `refreshToken`. The token is stored in memory and attached as `Authorization: Bearer <accessToken>` for all subsequent `/v1/...` calls.

### vCenter (optional but recommended)
- Uses **VCF.PowerCLI** (recommended) or **VMware.VimAutomation.Core**.  
- The script will connect with `Connect-VIServer` using the provided credentials to help verify inventory/mappings.

---

## Auto‑Population: Network Profile & IP Pools
When you click **Collect**, the tool runs the **Cluster Network Query** for `VCENTER_NSXT_NETWORK_CONFIG` and attempts to **auto‑fill** related fields (best‑effort):  
- **TEP IP Pool:** name, CIDR, gateway, range start/end  
- **Uplink Profile:** name, transport VLAN, teaming policy, active/standby uplinks  
- **vDS & Mapping:** NSX host switch `vdsName` and `vdsUplinkToNsxUplink` pairs  
These values are written into the hidden text fields used to build the JSON payload, and you can review/override before **Generate/Validate/Execute**.

> **Note:** In v0.10.33 the UI shows a **Name Suffix** field that is intended for future name derivation; the script primarily relies on **auto‑filled** values from the network query or the provided defaults (e.g., `az2-tep-pool`, `az2-uplink-profile`).

---

## Prerequisites
- **PowerShell 7+** and .NET/WPF available on the workstation.  
- **Modules:** `ImportExcel` (for workbook output), **VCF.PowerCLI** (recommended), **VMware.PowerCLI** (optional). The UI provides buttons to install them.
- **AZ2 hosts commissioned** in SDDC Manager and in **UNASSIGNED/USEABLE** state (so the Host Picker can find them).
- **Witness node** deployed and reachable; **Witness vSAN IP/CIDR** known; Witness added to vCenter. (See the Method of Procedure for detailed steps.)

---

## How to Run
1. Launch **PowerShell 7** and run the script: `pwsh.exe -File .\VCF9-StretchCluster-Automation.ps1`. The script **self‑signs** on first run and relaunches in **STA** as needed.
2. In **Prerequisites**, click **Recheck** (and **Install** buttons if needed).
3. Enter **SDDC Manager** FQDN/credentials and click **Connect**. Select the target **cluster**.
4. (Optional) Enter **vCenter** details and click **Connect/Verify**.
5. Click **Collect** to auto‑fill network fields; review and adjust **Network Profile**, **TEP pool**, and **uplink** details as needed.
6. Use **Pick UNASSIGNED/USEABLE Hosts** to populate AZ2 host FQDNs.
7. Provide **Witness** FQDN, vSAN IP, vSAN CIDR, and traffic sharing choice.
8. Click **Generate** to produce JSON payloads, **Validate** to POST validations, and **Execute** to PATCH the stretch. All responses are saved in the **Run** folder created under your selected output path.

---

## Outputs
- `clusterStretchSpec_*.json` and validation wrapper JSON written to the **Run** folder  
- Validation / execution API responses saved as timestamped JSON  
- Optional: Updated Excel workbook if a template is provided  
All actions are logged to `VCFStretch-*.log`.

---

## Troubleshooting
- **Connect failed / token issues**: Verify SDDC Manager FQDN and credentials; the tool uses `POST /v1/tokens` and reports HTTP codes/snippets for failures.
- **Auto‑fill empty**: The network query shape can vary by environment; review the saved `ClusterNetworkConfig-<clusterId>.json` and fill required fields manually.
- **Validation errors**: Open the saved response JSON to find schema or inventory mismatches (uplinks, VLAN, pool ranges). Adjust fields and re‑validate.

---

## Security Notes
- Access/refresh tokens are held **in memory** for the session; the script writes only request/response JSONs and logs to the Run folder.
- Certificates are **self‑signed and trusted locally** for code signing on first run to reduce execution policy friction.

---

## License
Internal use. Provide attribution if you reuse portions of the script or documentation.
