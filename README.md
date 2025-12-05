## Azure NSG Excel Template

### Overview

This repository contains an Excel-based template for modelling, generating, and maintaining Azure Network Security Group (NSG) rulesets.  
Instead of hand-writing PowerShell, Azure CLI, or ARM templates, you define NSG rules in a structured Excel workbook and let the built‑in VBA macros generate the corresponding configuration for you.

The goal is to give IT professionals a **single source of truth** for NSG policy, with repeatable, auditable deployments across environments.

### Key Benefits for IT Professionals

- **Centralised NSG policy**: Maintain all NSG rules for one or more environments in a single Excel file, rather than scattered PowerShell scripts or portal changes.
- **Human‑readable rules**: Rules are defined in simple, tabular form (rule number, name, source, destination, ports, protocol, direction, action, etc.), easy to review with peers and auditors.
- **Script generation, not manual typing**: The template generates configuration files for:
  - **PowerShell** (`Az` module)
  - **Azure CLI** (`az network nsg …`)
  - **ARM template JSON** (for ARM deployments)
  - **YAML** (for documentation or pipeline consumption)
- **Idempotent rule management**: Generated scripts:
  - Ensure the target NSG exists (and create it if not).
  - Track the rules defined in the spreadsheet and **remove rules that no longer exist in Excel**, keeping the NSG aligned with the approved design.
- **Application Security Group (ASG) support**: A dedicated sheet allows defining ASGs once; the generator then creates the ASG resources (PowerShell, AZ CLI, or ARM) and references them from NSG rules.
- **Environment metadata baked in**: NSG name, resource group, Azure region, “local subnets”, and other metadata are captured in the workbook and injected automatically into generated scripts.
- **Governance‑friendly**: Because the Excel file is small and text‑like in structure, it is easy to version with Git, attach to change requests, and use as evidence for firewall/NSG reviews.

### How It Works (High Level)

- The Excel workbook contains:
  - A **Config** sheet that controls generation options (output format, target OS, output file name and directory, etc.).
  - One or more **NSG rule sheets** where you define inbound/outbound rules in rows.
  - Supporting sheets such as **ASGs** are used by the macros to map columns to PowerShell/AZ CLI/ARM/YAML syntax.
  - Depending on the selected format (`Powershell`, `AZCLI`, `ARM`, or `YAML`), the macro prints the appropriate syntax for each rule and any associated ASGs.

### Supported Outputs

- **PowerShell (`Az` module)**  
  - Looks up the NSG in the target resource group, creates it if missing.  
  - Creates or updates security rules according to the Excel definition.  
  - Identifies rules in Azure that are not present in the spreadsheet and removes them.

- **Azure CLI**  
  - Supports both **Windows** and **Linux/macOS** shell styles.  
  - Ensures the NSG exists, creates it if it does not.  
  - Tracks the set of desired rule names (from Excel) and deletes any extra rules from the NSG in Azure.

- **ARM Template JSON**  
  - Builds a JSON object with NSG properties and `securityRules` based on the rows in the Excel sheet.  
  - Suitable for use in ARM deployments or as a starting point for Bicep.

- **YAML**  
  - Outputs a YAML representation of the NSG ruleset, useful for documentation, code reviews, or CI/CD pipelines that ingest YAML.

### Typical Workflow for Managing NSGs

1. **Clone or download** this repository and open the Excel NSG template.
2. On the **Config** sheet:
   - Set `Directory` and `Filename` to where the generated file should be written.
   - Set `Format` to your desired output (`Powershell`, `AZCLI`, `ARM`, or `YAML`).
   - Optionally set `Operating System` (e.g. `Windows` or `Linux`) to control shell‑specific CLI syntax.
3. (Optional) On the **ASGs** sheet:
   - Define Application Security Groups and their mappings.
4. On the **NSG rule sheet**:
   - Enter **NSG Name**, **Resource Group**, **Local Subnets** and **Location**.
   - In the rules section (rows with numeric IDs in column B), define each rule’s:
     - Name / description
     - Priority
     - Source / destination (IP prefixes, ASGs, or “Local Subnets”)
     - Protocol and ports
     - Direction and action (Allow/Deny)
   - Note that the rules are broken into blocks to help with organisation and future development of the ruleset
5. Run the **`Generate NSG Rules`** macro from the **NSG rule sheet** tab:
   - Excel generates the output file in the configured directory.
6. **Review** the generated script or template into your standard change‑control process:
   - Attach to change tickets or CAB documentation.
7. **Deploy**:
   - For PowerShell/AZ CLI: run the generated script in your preferred automation context (local shell, automation account, DevOps pipeline, etc.).
   - For ARM: deploy via `New-AzResourceGroupDeployment`, `az deployment group create`, or your chosen pipeline.

### Prerequisites

- **Excel for Windows** with macros enabled (the solution relies on VBA modules).
- For deployment, depending on the chosen output format:
  - **PowerShell** with the **Az** module.
  - **Azure CLI** (`az`) installed and authenticated.
  - Ability to deploy **ARM templates** if using the ARM output.

### Licensing

- This template is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0).
- You may use and modify it freely for personal or internal business use, but you may not sell it or use it commercially.
- **Free version (this workbook)**: Free to download and use, but technically limited to generating configuration for up to **5 NSG rules** per NSG. This is intended as a starter / evaluation version of the tool.
- **Pro version**: A separate paid edition of the workbook removes the 5‑rule limit and may include additional productivity features. The Pro version is distributed under a separate commercial licence.

### Disclaimer

This template is intended to assist with **modelling and automating NSG configuration**, but it does not replace a proper design or security review.  
Always validate generated scripts in a non‑production environment and follow your organisation’s change‑management and security governance processes.


