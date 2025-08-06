<div align="center">
  <img src="imgs/py-powershell.png" alt="py-powershell logo" height="300" />

  <h1>py-powershell</h1>
  
  <p><em>Python module to manage persistent PowerShell sessions, with support for command execution, JSON result parsing, and secure error handling.</em></p>


---

**py-powershell** is a modern, lightweight Python wrapper that provides seamless integration with PowerShell sessions. Built for developers who need reliable, persistent PowerShell automation in their Python applications.

## Key Features

- **Persistent Sessions** - Maintain long-running PowerShell sessions for faster execution and stateful operations
- **Persistent Authentication** - Stay logged in to PowerShell modules and tools across multiple commands
- **Secure by Design** - Built-in input sanitization prevents command injection attacks
- **Asynchronous Processing** - Non-blocking command execution with configurable timeout handling
- **Clean Output** - Automatic removal of ANSI escape sequences for parseable results
- **JSON Support** - Native JSON parsing from PowerShell `ConvertTo-Json` output
- **Rich Logging** - Automatic error logging from PowerShell stderr during command execution

## Architecture

The `PowerShellSession` class provides the core functionality:

- **Subprocess Management**: Manages a persistent PowerShell process with stdin/stdout/stderr pipes
- **Thread Safety**: Uses threading for non-blocking I/O operations
- **Security**: Input sanitization prevents command injection attacks
- **Output Processing**: Automatic ANSI escape sequence removal and JSON parsing
- **Error Handling**: Automatic logging of PowerShell stderr output during each command execution

### Class Hierarchy

```
PowerShellSession
├── __init__()          # Initialize persistent session
├── execute()           # Execute commands with optional JSON parsing
├── read_output()       # Read and process command output
├── json_parse_output() # Parse JSON from PowerShell output
├── sanitize()          # Sanitize input to prevent injection
├── remove_ansi()       # Clean ANSI escape sequences
└── close()             # Cleanup and terminate session
```

## Requirements

- **Python 3.9+** - Modern Python with full type hint support
- **PowerShell Core (pwsh)** - Must be installed and available in PATH
- **Operating System**: Windows, macOS, or Linux with PowerShell Core

### Installing PowerShell Core

**Windows:**
```powershell
winget install Microsoft.PowerShell
```

**macOS:**
```bash
brew install powershell
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install -y powershell
```

## Quick Start

### Installation

```bash
pip install py-powershell
```

### Basic Usage

```python
from py_powershell import PowerShellSession

# Create a persistent PowerShell session
pwsh = PowerShellSession()

# Execute a simple command
result = pwsh.execute("Get-Process | Select-Object -First 5")
print(result)

# Execute with JSON parsing and custom timeout
data = pwsh.execute("Get-Service | Select-Object -First 3", json_parse=True, timeout=15)
print(data)

# Always close the session when done
pwsh.close()
```

### Context Manager

```python
from py_powershell import PowerShellSession

# Use context manager for automatic cleanup
with PowerShellSession() as pwsh:
    # Execute commands with different timeouts as needed
    services = pwsh.execute("Get-Service", json_parse=True, timeout=20)
    
    # Work with the results
    for service in services:
        print(f"Service: {service['Name']} - Status: {service['Status']}")
```

## Comprehensive Examples

### Working with Azure PowerShell

```python
from py_powershell import PowerShellSession
import logging

# Configure logging to see PowerShell errors automatically captured
logging.basicConfig(level=logging.INFO)

with PowerShellSession() as pwsh:
    # Connect to Azure (interactive login)
    pwsh.execute("Connect-AzAccount")
    
    # Get Azure VMs with JSON output
    # Any PowerShell errors will be automatically logged by the session
    vms = pwsh.execute(
        "Get-AzVM | Select-Object Name, ResourceGroupName, Location",
        json_parse=True,
        timeout=30
    )
    
    # Process results
    for vm in vms:
        print(f"VM: {vm['Name']} in {vm['ResourceGroupName']} ({vm['Location']})")
```

### Microsoft 365 Administration

```python
from py_powershell import PowerShellSession

with PowerShellSession() as pwsh:
    # Connect to Exchange Online
    pwsh.execute("Connect-ExchangeOnline")
    
    # Get mailbox statistics
    mailboxes = pwsh.execute(
        "Get-Mailbox -ResultSize 10 | Get-MailboxStatistics",
        json_parse=True,
        timeout=60
    )
    
    # Display mailbox info
    for mailbox in mailboxes:
        size_mb = float(mailbox['TotalItemSize'].split('(')[1].split(' ')[0].replace(',', '')) / 1024 / 1024
        print(f"Mailbox: {mailbox['DisplayName']} - Size: {size_mb:.2f} MB")
```

### System Administration

```python
from py_powershell import PowerShellSession

def get_system_info():
    with PowerShellSession() as pwsh:
        # Get computer info
        computer_info = pwsh.execute(
            "Get-ComputerInfo | Select-Object WindowsProductName, TotalPhysicalMemory, CsProcessors",
            json_parse=True
        )
        
        # Get disk usage
        disk_info = pwsh.execute(
            "Get-WmiObject -Class Win32_LogicalDisk | Select-Object DeviceID, Size, FreeSpace",
            json_parse=True
        )
        
        return {
            'computer': computer_info,
            'disks': disk_info
        }

system_data = get_system_info()
print(f"OS: {system_data['computer']['WindowsProductName']}")
for disk in system_data['disks']:
    free_gb = int(disk['FreeSpace']) / (1024**3)
    total_gb = int(disk['Size']) / (1024**3)
    print(f"Drive {disk['DeviceID']} - {free_gb:.1f} GB free of {total_gb:.1f} GB")
```

### Error Handling and Logging

```python
import logging
from py_powershell import PowerShellSession

# Configure logging to see PowerShell errors
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

with PowerShellSession() as pwsh:
    # PowerShell errors are automatically logged during execute()
    # The logger captures stderr output from PowerShell
    result = pwsh.execute("Get-NonExistentCommand", timeout=5)
    
    # Any PowerShell errors will appear in logs like:
    # ERROR - py_powershell.powershell_session - PowerShell error output: Get-NonExistentCommand : The term 'Get-NonExistentCommand' is not recognized...
    
    if not result:
        print("Command returned no output or failed")
```

## Advanced Configuration

### Custom Timeout Configuration

```python
from py_powershell import PowerShellSession

with PowerShellSession() as pwsh:
    # Execute with custom timeout (default is 10 seconds)
    result = pwsh.execute("Get-Process", timeout=30)
    
    # Long-running command with extended timeout
    large_data = pwsh.execute(
        "Get-EventLog -LogName System -Newest 1000",
        json_parse=True,
        timeout=60
    )
    
    # Quick command with short timeout
    quick_result = pwsh.execute("Get-Date", timeout=5)
    
    # Very long operation (e.g., Azure operations)
    azure_vms = pwsh.execute(
        "Get-AzVM -Status",
        json_parse=True,
        timeout=120  # 2 minutes
    )
```

### Secure Credential Handling

```python
from py_powershell import PowerShellSession
import os

def connect_with_credentials():
    with PowerShellSession() as pwsh:
        # Use environment variables for credentials
        username = pwsh.sanitize(os.environ.get('pwsh_USERNAME', ''))
        password = pwsh.sanitize(os.environ.get('pwsh_PASSWORD', ''))
        
        if username and password:
            # Create secure credential
            pwsh.execute(f"$secpass = ConvertTo-SecureString '{password}' -AsPlainText -Force")
            pwsh.execute(f"$cred = New-Object System.Management.Automation.pwshCredential('{username}', $secpass)")
            
            # Use credential for authentication
            result = pwsh.execute("Connect-SomeService -Credential $cred")
            return result
        else:
            raise ValueError("Credentials not provided in environment variables")
```

## Contributing

We welcome contributions! Please submit pull requests or open issues on GitHub.

## Support

- **Documentation**: [Read the full documentation](httpwsh://py-powershell.readthedocs.io)
- **Issues**: [Report bugs or request features](httpwsh://github.com/hugopbrito/py-powershell/issues)
- **Discussions**: [Join the community discussions](httpwsh://github.com/hugopbrito/py-powershell/discussions)
