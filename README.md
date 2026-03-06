Helpdesk Companion

Helpdesk Companion is a professional PowerShell-based WPF application designed to streamline help desk operations, improve response times, and simplify remote troubleshooting. It provides a centralized interface for Active Directory management, remote system diagnostics, and ITIL service desk integration.

🚀 Key Features

Active Directory Management:

Real-time detection and logging of account lockouts.

One-click account unlocking with de-duplicated logging logic.

Secure complex password generation and resets.

Group membership visualization.

Lockout source finder (automatically queries the PDC for Event ID 4740).

Remote System Diagnostics:

Remote process and service management.

System information and hardware device inventory.

Uptime checking with connection timeout protection.

User profile management (listing and deletion).

Silent uninstallation of remote software.

Remote power actions (Restart/Shutdown).

Service Desk Integration (Freshservice):

Instant ticket creation ("Quick Ticket") with custom fields.

User asset lookup to identify assigned hardware.

Deep-linking to Freshservice requester and asset records.

Professional UI/UX:

Modern Fluent-inspired WPF interface.

Native Light and Dark mode support.

Modeless tool windows for multitasking.

Responsive de-duplicated logging pane.

📂 Project Structure

HelpdeskCompanion/
├── Main.ps1                     # Application entry point
├── PrinterManager.ps1           # Standalone printer management logic
├── hdcompanioncfg.json          # Configuration settings
├── modules/
│   ├── UIEvents.psm1            # Master Router (delegates events)
│   ├── UI.Core.psm1             # Base UI logic (Theming, Search, Logs)
│   ├── UI.UserActions.psm1      # AD User context menus and overlays
│   ├── UI.ComputerActions.psm1  # AD Computer tools and system manager
│   ├── UI.Freshservice.psm1     # Freshservice API and UI logic
│   ├── ActiveDirectory.psm1     # Backend AD query logic
│   ├── CoreLogic.psm1           # App config and state management
│   └── RemoteManagement.psm1    # WMI/CIM remote execution logic
└── UI/                          # XAML layouts
    ├── MainWindow.xaml
    ├── Windows/                 # Secondary tool windows
    └── Dialogs/                 # Custom message boxes and prompts


🛠️ Requirements

OS: Windows 10/11 or Windows Server 2016+

PowerShell: Windows PowerShell 5.1

Modules: ActiveDirectory RSAT tools must be installed.

Permissions: Admin rights on remote computers (for WMI/CIM) and AD delegated permissions for account management.

⚙️ Configuration

The application primarily loads configuration from a central UNC path (defined in CoreLogic.psm1) with a local fallback to hdcompanioncfg.json.

Key settings include:

DomainName: Your target AD domain.

LogDirectoryUNC: Shared path for centralized CSV logging.

FreshserviceAPIKey: Your API key for ticket/asset integration.

🔧 Installation & Usage

Clone or copy the folder structure to your workstation or a network share.

Ensure you have the Active Directory RSAT modules installed.

Execute the application via Main.ps1:

pwsh.exe -File .\Main.ps1


Use the Search bar to find users or computers.

Right-click any entry in the main list to access specialized tool menus.

🛡️ Architecture Notes

Modularity: UI events are decoupled into domain modules to prevent "God Function" complexity.

Resiliency: Remote queries are wrapped in background jobs with manual timeout loops to prevent the UI from freezing on unresponsive or firewalled endpoints.

State Management: Uses a global $State object to synchronize data across modules.
