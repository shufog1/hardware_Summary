# Define global variables to store key information as we collect it
$global:ProcessorInfo = "Unknown"
$global:StorageInfo = "Unknown"
$global:GraphicsInfo = "Unknown"
$global:RamInfo = "Unknown"
$global:SlotInfo = "Unknown"
$global:DiskCount = 0

# Set error action preference
$ErrorActionPreference = "SilentlyContinue"

# Create timestamp for the filename
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputFile = "$env:USERPROFILE\Desktop\HardwareInventory-$timestamp.txt"

# Function to write section headers
function Write-SectionHeader {
    param ([string]$Title)
    
    $header = "`n" + "="*80 + "`n" + " $Title " + "`n" + "="*80
    Add-Content -Path $outputFile -Value $header
}

# Function to handle errors
function Write-ErrorInfo {
    param (
        [string]$Component,
        [string]$ErrorMsg
    )
    
    Add-Content -Path $outputFile -Value "`n[ERROR] Failed to retrieve $Component information: $ErrorMsg`n"
}

# Initialize the output file with a header
$computerName = $env:COMPUTERNAME
"HARDWARE INVENTORY REPORT" | Out-File -FilePath $outputFile
"Computer Name: $computerName" | Add-Content -Path $outputFile
"Generated on: $(Get-Date)" | Add-Content -Path $outputFile
"="*80 | Add-Content -Path $outputFile

try {
    # System Information
    Write-SectionHeader "SYSTEM INFORMATION"
    
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        $biosInfo = Get-WmiObject -Class Win32_BIOS
        
        "Computer Manufacturer: $($computerSystem.Manufacturer)" | Add-Content -Path $outputFile
        "Computer Model: $($computerSystem.Model)" | Add-Content -Path $outputFile
        "Serial Number: $($biosInfo.SerialNumber)" | Add-Content -Path $outputFile
    }
    catch {
        Write-ErrorInfo -Component "System Information" -ErrorMsg $_.Exception.Message
    }
    
    # Processor Information
    Write-SectionHeader "PROCESSOR INFORMATION"
    
    try {
        $processors = Get-WmiObject -Class Win32_Processor
        
        foreach ($processor in $processors) {
            # Store the processor name in our global variable
            $global:ProcessorInfo = $processor.Name
            
            "Processor: $($processor.Name)" | Add-Content -Path $outputFile
            "Cores: $($processor.NumberOfCores)" | Add-Content -Path $outputFile
            "Logical Processors: $($processor.NumberOfLogicalProcessors)" | Add-Content -Path $outputFile
            "Max Clock Speed: $($processor.MaxClockSpeed) MHz" | Add-Content -Path $outputFile
        }
    }
    catch {
        Write-ErrorInfo -Component "Processor Information" -ErrorMsg $_.Exception.Message
    }
    
    # Memory Slots and Information
    Write-SectionHeader "MEMORY INFORMATION"
    
    try {
        # Direct query for memory modules
        $physicalMemory = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue
        $memorySlots = Get-WmiObject -Class Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue
        
        # Count how many actual modules we have
        $moduleCount = 0
        if ($physicalMemory -ne $null) {
            if ($physicalMemory -is [array]) {
                $moduleCount = $physicalMemory.Count
            } else {
                $moduleCount = 1  # Single module
            }
        }
        
        # Get total slots from memory array
        $totalSlots = 0
        if ($memorySlots -ne $null) {
            if ($memorySlots -is [array]) {
                foreach ($slot in $memorySlots) {
                    $totalSlots += $slot.MemoryDevices
                }
            } else {
                $totalSlots = $memorySlots.MemoryDevices
            }
        }
        
        # Make sure total slots is at least equal to module count
        if ($totalSlots -lt $moduleCount) {
            $totalSlots = $moduleCount
        }
        
        # Calculate empty slots
        $emptySlots = $totalSlots - $moduleCount
        
        # Store RAM and slot info in global variables
        $global:SlotInfo = "$moduleCount used of $totalSlots total"
        
        # Write information to file with explicit string values
        "Total Memory Slots: $totalSlots" | Add-Content -Path $outputFile
        "Used Memory Slots: $moduleCount" | Add-Content -Path $outputFile
        "Empty Memory Slots: $emptySlots" | Add-Content -Path $outputFile
        
        # Process each memory module
        $totalRam = 0
        $i = 0
        
        # Handle both array and single module cases
        if ($physicalMemory -ne $null) {
            if ($physicalMemory -is [array]) {
                # Multiple modules
                foreach ($module in $physicalMemory) {
                    $sizeGB = [math]::Round($module.Capacity / 1GB, 2)
                    $totalRam += $sizeGB
                    $speed = $module.Speed
                    
                    "Slot $i Details:" | Add-Content -Path $outputFile
                    "  Manufacturer: $($module.Manufacturer)" | Add-Content -Path $outputFile
                    "  Part Number: $($module.PartNumber.Trim())" | Add-Content -Path $outputFile
                    "  Capacity: $sizeGB GB" | Add-Content -Path $outputFile
                    "  Speed: $speed MHz" | Add-Content -Path $outputFile
                    "  Form Factor: $($module.FormFactor)" | Add-Content -Path $outputFile
                    "  Memory Type: $($module.MemoryType)" | Add-Content -Path $outputFile
                    "  Bank/Device Locator: $($module.BankLabel)/$($module.DeviceLocator)" | Add-Content -Path $outputFile
                    $i++
                }
            } else {
                # Single module
                $sizeGB = [math]::Round($physicalMemory.Capacity / 1GB, 2)
                $totalRam += $sizeGB
                $speed = $physicalMemory.Speed
                
                "Slot $i Details:" | Add-Content -Path $outputFile
                "  Manufacturer: $($physicalMemory.Manufacturer)" | Add-Content -Path $outputFile
                "  Part Number: $($physicalMemory.PartNumber.Trim())" | Add-Content -Path $outputFile
                "  Capacity: $sizeGB GB" | Add-Content -Path $outputFile
                "  Speed: $speed MHz" | Add-Content -Path $outputFile
                "  Form Factor: $($physicalMemory.FormFactor)" | Add-Content -Path $outputFile
                "  Memory Type: $($physicalMemory.MemoryType)" | Add-Content -Path $outputFile
                "  Bank/Device Locator: $($physicalMemory.BankLabel)/$($physicalMemory.DeviceLocator)" | Add-Content -Path $outputFile
            }
        } else {
            "No memory modules detected" | Add-Content -Path $outputFile
            $totalRam = 0
        }
        
        # Store RAM info in global variable
        $global:RamInfo = "$totalRam GB"
        
        "Total Installed RAM: $totalRam GB" | Add-Content -Path $outputFile
    }
    catch {
        Write-ErrorInfo -Component "Memory Information" -ErrorMsg $_.Exception.Message
    }
    
    # Storage Information
    Write-SectionHeader "STORAGE INFORMATION"
    
    try {
        # Get physical disks directly with error handling
        $diskDrives = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue
        
        # Set default storage info in case the detection fails
        $global:StorageInfo = "Unknown"
        
        if (-not $diskDrives -or $diskDrives.Count -eq 0) {
            "No physical disks detected" | Add-Content -Path $outputFile
            $global:StorageInfo = "No drives detected"
        } else {
            # This will work even if $diskDrives is not an array but a single object
            $diskCount = if ($diskDrives -is [array]) { $diskDrives.Count } else { 1 }
            
            # Explicitly set count for summary and details
            "Total Physical Disks: $diskCount" | Add-Content -Path $outputFile
            $global:StorageInfo = "$diskCount drive"
            if ($diskCount -ne 1) { $global:StorageInfo += "s" }
            
            # Process each physical disk
            foreach ($disk in $diskDrives) {
                $diskModel = $disk.Model
                $diskSize = [math]::Round($disk.Size / 1GB, 2)
                $diskInterface = $disk.InterfaceType
                
                # Determine connection type with fallback logic
                $connectionType = "Unknown"
                if ($disk.PNPDeviceID -like "*NVME*" -or $diskModel -like "*NVME*") {
                    $connectionType = "M.2 NVMe"
                } elseif ($diskInterface -eq "SCSI" -and ($diskModel -like "*SSD*" -or $diskModel -like "*M.2*")) {
                    $connectionType = "M.2 SATA"
                } elseif ($diskInterface -eq "SCSI" -or $diskInterface -eq "IDE") {
                    $connectionType = "SATA"
                }
                
                # Determine disk type based on model name if we can't get it from WMI
                $diskType = "Unknown"
                if ($diskModel -like "*SSD*" -or $connectionType -like "*M.2*") {
                    $diskType = "SSD"
                } elseif ($diskModel -like "*HDD*") {
                    $diskType = "HDD"
                }
                
                # Output disk information
                "Disk: $diskModel" | Add-Content -Path $outputFile
                "  Type: $diskType" | Add-Content -Path $outputFile
                "  Connection: $connectionType" | Add-Content -Path $outputFile
                "  Interface: $diskInterface" | Add-Content -Path $outputFile
                "  Size: $diskSize GB" | Add-Content -Path $outputFile
                
                # Get partitions with explicit error handling and count
                $partitions = Get-WmiObject -Class Win32_DiskPartition -Filter "DiskIndex=$($disk.Index)" -ErrorAction SilentlyContinue
                $partitionCount = if ($partitions -is [array]) { $partitions.Count } else { if ($partitions) { 1 } else { 0 } }
                
                "  Partitions: $partitionCount" | Add-Content -Path $outputFile
                if ($partitionCount -gt 0) {
                    foreach ($partition in $partitions) {
                        $partitionSize = [math]::Round($partition.Size / 1GB, 2)
                        "    - $($partition.Name): $partitionSize GB" | Add-Content -Path $outputFile
                    }
                } else {
                    "    - None detected" | Add-Content -Path $outputFile
                }
                
                # Add detailed info to global variable for summary
                $global:StorageInfo = "$diskCount drive - $diskModel ($connectionType, $diskType, $diskSize GB)"
            }
        }
    }
    catch {
        Write-ErrorInfo -Component "Storage Information" -ErrorMsg $_.Exception.Message
        # Set a fallback for summary
        $global:StorageInfo = "Detection failed"
    }
    
    # Expansion Slots
    Write-SectionHeader "EXPANSION SLOTS"
    
    try {
        $expansionSlots = Get-WmiObject -Class Win32_SystemSlot
        
        "Total Expansion Slots: $($expansionSlots.Count)" | Add-Content -Path $outputFile
        
        foreach ($slot in $expansionSlots) {
            $status = if ($slot.Status -eq "OK" -and $slot.CurrentUsage -eq "Available") { "Empty" } else { "In Use" }
            
            "Slot: $($slot.SlotDesignation)" | Add-Content -Path $outputFile
            "  Description: $($slot.Description)" | Add-Content -Path $outputFile
            "  Type: $($slot.Name)" | Add-Content -Path $outputFile
            "  Status: $status" | Add-Content -Path $outputFile
        }
    }
    catch {
        Write-ErrorInfo -Component "Expansion Slots" -ErrorMsg $_.Exception.Message
    }
    
    # Graphics Card Information
    Write-SectionHeader "GRAPHICS CARD INFORMATION"
    
    try {
        # Get all video controllers
        $graphicsCards = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue
        
        # Filter out virtual display adapters
        $realGraphicsCards = $graphicsCards | Where-Object {
            $_.Name -notlike "*Microsoft*Basic*Display*" -and
            $_.Name -notlike "*Microsoft*Remote*Display*" -and
            $_.Name -notlike "*MS Idd*"
        }
        
        if (-not $realGraphicsCards -or $realGraphicsCards.Count -eq 0) {
            # If filtering removed all cards, use original list
            $displayGraphicsCards = $graphicsCards
        } else {
            $displayGraphicsCards = $realGraphicsCards
        }
        
        # Categorize GPUs
        $integratedGPUs = $displayGraphicsCards | Where-Object {
            $_.Name -like "*Intel*HD*" -or
            $_.Name -like "*Intel*UHD*" -or
            $_.Name -like "*Intel*Iris*" -or
            $_.Name -like "*AMD*Radeon*Graphics*" -or # APU pattern
            $_.Name -like "*AMD*Graphics*"
        }
        
        $discreteGPUs = $displayGraphicsCards | Where-Object {
            ($_.Name -like "*NVIDIA*" -or
            $_.Name -like "*GeForce*" -or
            $_.Name -like "*Quadro*" -or
            $_.Name -like "*RTX*" -or
            $_.Name -like "*GTX*" -or
            $_.Name -like "*Radeon*" -or
            $_.Name -like "*AMD*") -and
            $_.Name -notlike "*Intel*HD*" -and
            $_.Name -notlike "*Intel*UHD*" -and
            $_.Name -notlike "*Intel*Iris*" -and
            $_.Name -notlike "*AMD*Graphics*"
        }
        
        # Store first real GPU in global variable for summary
        $global:GraphicsInfo = "Unknown"
        
        if ($integratedGPUs -and $integratedGPUs.Count -gt 0) {
            $global:GraphicsInfo = $integratedGPUs[0].Name
        }
        
        if ($discreteGPUs -and $discreteGPUs.Count -gt 0) {
            if ($global:GraphicsInfo -ne "Unknown") {
                # We have both integrated and discrete
                $global:GraphicsInfo += " + $($discreteGPUs[0].Name)"
            } else {
                # Only discrete
                $global:GraphicsInfo = $discreteGPUs[0].Name
            }
        }
        
        # If we still don't have a GPU name, try to get any valid GPU
        if ($global:GraphicsInfo -eq "Unknown" -and $displayGraphicsCards -and $displayGraphicsCards.Count -gt 0) {
            $global:GraphicsInfo = $displayGraphicsCards[0].Name
        }
        
        # Report integrated GPUs
        if ($integratedGPUs -and $integratedGPUs.Count -gt 0) {
            "Integrated Graphics:" | Add-Content -Path $outputFile
            foreach ($gpu in $integratedGPUs) {
                $vramMB = if($gpu.AdapterRAM -gt 0) { [math]::Round($gpu.AdapterRAM / 1MB, 2) } else { "Shared" }
                
                "  GPU: $($gpu.Name)" | Add-Content -Path $outputFile
                "    Type: Integrated" | Add-Content -Path $outputFile
                "    Video Processor: $($gpu.VideoProcessor)" | Add-Content -Path $outputFile
                "    Driver Version: $($gpu.DriverVersion)" | Add-Content -Path $outputFile
                "    Video Memory: $vramMB MB" | Add-Content -Path $outputFile
                if ($gpu.CurrentHorizontalResolution -gt 0) {
                    "    Current Resolution: $($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution)" | Add-Content -Path $outputFile
                    "    Refresh Rate: $($gpu.CurrentRefreshRate) Hz" | Add-Content -Path $outputFile
                }
            }
        }
        
        # Report discrete GPUs
        if ($discreteGPUs -and $discreteGPUs.Count -gt 0) {
            "Discrete Graphics:" | Add-Content -Path $outputFile
            foreach ($gpu in $discreteGPUs) {
                $vramMB = if($gpu.AdapterRAM -gt 0) { [math]::Round($gpu.AdapterRAM / 1MB, 2) } else { "Unknown" }
                
                "  GPU: $($gpu.Name)" | Add-Content -Path $outputFile
                "    Type: Discrete" | Add-Content -Path $outputFile
                "    Video Processor: $($gpu.VideoProcessor)" | Add-Content -Path $outputFile
                "    Driver Version: $($gpu.DriverVersion)" | Add-Content -Path $outputFile
                "    Video Memory: $vramMB MB" | Add-Content -Path $outputFile
                if ($gpu.CurrentHorizontalResolution -gt 0) {
                    "    Current Resolution: $($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution)" | Add-Content -Path $outputFile
                    "    Refresh Rate: $($gpu.CurrentRefreshRate) Hz" | Add-Content -Path $outputFile
                }
            }
        }
        
        # If no categorized GPUs, show all GPUs
        if ((-not $integratedGPUs -or $integratedGPUs.Count -eq 0) -and 
            (-not $discreteGPUs -or $discreteGPUs.Count -eq 0) -and 
            $displayGraphicsCards -and $displayGraphicsCards.Count -gt 0) {
            
            "Graphics Devices:" | Add-Content -Path $outputFile
            foreach ($gpu in $displayGraphicsCards) {
                $vramMB = if($gpu.AdapterRAM -gt 0) { [math]::Round($gpu.AdapterRAM / 1MB, 2) } else { "Unknown" }
                
                "  GPU: $($gpu.Name)" | Add-Content -Path $outputFile
                if ($gpu.VideoProcessor) { "    Video Processor: $($gpu.VideoProcessor)" | Add-Content -Path $outputFile }
                "    Driver Version: $($gpu.DriverVersion)" | Add-Content -Path $outputFile
                "    Video Memory: $vramMB MB" | Add-Content -Path $outputFile
                if ($gpu.CurrentHorizontalResolution -gt 0) {
                    "    Current Resolution: $($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution)" | Add-Content -Path $outputFile
                    "    Refresh Rate: $($gpu.CurrentRefreshRate) Hz" | Add-Content -Path $outputFile
                }
            }
        }
    }
    catch {
        Write-ErrorInfo -Component "Graphics Card Information" -ErrorMsg $_.Exception.Message
        $global:GraphicsInfo = "Detection failed"
    }
    
    # Network Adapters
    Write-SectionHeader "NETWORK ADAPTERS"
    
    try {
        $networkAdapters = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
        
        foreach ($adapter in $networkAdapters) {
            $adapterConfig = $adapter | Get-WmiObject -Class Win32_NetworkAdapterConfiguration
            
            "Network Adapter: $($adapter.Name)" | Add-Content -Path $outputFile
            "  Manufacturer: $($adapter.Manufacturer)" | Add-Content -Path $outputFile
            "  MAC Address: $($adapterConfig.MACAddress)" | Add-Content -Path $outputFile
            "  Connection Status: $($adapter.NetConnectionStatus)" | Add-Content -Path $outputFile
            
            if ($adapter.Name -match "Bluetooth" -or $adapter.Description -match "Bluetooth") {
                "  Type: Bluetooth" | Add-Content -Path $outputFile
            }
            elseif ($adapter.Name -match "Wi-?Fi" -or $adapter.Description -match "Wi-?Fi" -or $adapter.Name -match "Wireless" -or $adapter.Description -match "Wireless") {
                "  Type: Wireless" | Add-Content -Path $outputFile
                
                # Try to get wireless adapter properties
                try {
                    $wirelessAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -eq $adapter.Description }
                    
                    if ($wirelessAdapter) {
                        "  Speed: $([math]::Round($wirelessAdapter.LinkSpeed / 1000000, 0)) Mbps" | Add-Content -Path $outputFile
                    }
                }
                catch {
                    # Skip if we can't get wireless properties
                }
            }
            else {
                "  Type: Wired" | Add-Content -Path $outputFile
                
                # Try to get wired adapter speed
                try {
                    $wiredAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -eq $adapter.Description }
                    
                    if ($wiredAdapter) {
                        "  Speed: $([math]::Round($wiredAdapter.LinkSpeed / 1000000, 0)) Mbps" | Add-Content -Path $outputFile
                    }
                }
                catch {
                    # Skip if we can't get speed
                }
            }
        }
    }
    catch {
        Write-ErrorInfo -Component "Network Adapters" -ErrorMsg $_.Exception.Message
    }
    
    # Bluetooth Devices (simple detection only)
    Write-SectionHeader "BLUETOOTH INFORMATION"
    
    try {
        # Multiple detection methods for Bluetooth
        $bluetoothRadios = Get-WmiObject -Namespace "root\WMI" -Class "BthRadio" -ErrorAction SilentlyContinue
        $bluetoothInterfaces = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Bluetooth*" } -ErrorAction SilentlyContinue
        $bluetoothDevices = Get-WmiObject -Class Win32_PnPEntity | Where-Object { 
            $_.Name -like "*Bluetooth*" -or 
            $_.Description -like "*Bluetooth*" -or 
            $_.Service -like "*BTHUSB*" -or 
            $_.Service -like "*BTHMINI*" 
        } -ErrorAction SilentlyContinue
        
        # Check if any detection method found Bluetooth
        $global:bluetoothFound = ($bluetoothRadios -and $bluetoothRadios.Count -gt 0) -or
                         ($bluetoothInterfaces -and $bluetoothInterfaces.Count -gt 0) -or
                         ($bluetoothDevices -and $bluetoothDevices.Count -gt 0)
        
        if ($global:bluetoothFound) {
            "Bluetooth Detected: Yes" | Add-Content -Path $outputFile
            
            # Simple detection information
            "  (Bluetooth hardware was detected using one or more detection methods)" | Add-Content -Path $outputFile
        }
        else {
            "Bluetooth Detected: No" | Add-Content -Path $outputFile
            "  (Note: Detection may fail even if Bluetooth exists - check Device Manager)" | Add-Content -Path $outputFile
            
            # Try registry check as a last resort
            try {
                $bluetoothRegistry = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\BTHUSB" -ErrorAction SilentlyContinue
                
                if ($bluetoothRegistry) {
                    "  Bluetooth Registry Entry Found: Yes" | Add-Content -Path $outputFile
                    "  Bluetooth may be installed but not detected by PowerShell" | Add-Content -Path $outputFile
                    $global:bluetoothFound = $true
                }
            }
            catch {
                # Skip if registry check fails
            }
        }
    }
    catch {
        Write-ErrorInfo -Component "Bluetooth Information" -ErrorMsg $_.Exception.Message
    }
    
    # Sound Devices
    Write-SectionHeader "SOUND DEVICES"
    
    try {
        $soundDevices = Get-WmiObject -Class Win32_SoundDevice
        
        foreach ($device in $soundDevices) {
            "Sound Device: $($device.Name)" | Add-Content -Path $outputFile
            "  Manufacturer: $($device.Manufacturer)" | Add-Content -Path $outputFile
            "  Status: $($device.Status)" | Add-Content -Path $outputFile
        }
    }
    catch {
        Write-ErrorInfo -Component "Sound Devices" -ErrorMsg $_.Exception.Message
    }
    
    # BIOS Information
    Write-SectionHeader "BIOS INFORMATION"
    
    try {
        $biosInfo = Get-WmiObject -Class Win32_BIOS
        
        "BIOS Manufacturer: $($biosInfo.Manufacturer)" | Add-Content -Path $outputFile
        "BIOS Version: $($biosInfo.SMBIOSBIOSVersion)" | Add-Content -Path $outputFile
        "BIOS Release Date: $($biosInfo.ReleaseDate)" | Add-Content -Path $outputFile
    }
    catch {
        Write-ErrorInfo -Component "BIOS Information" -ErrorMsg $_.Exception.Message
    }
    
    # Summary
    Write-SectionHeader "HARDWARE SUMMARY"
    
    try {
        "Computer Name: $computerName" | Add-Content -Path $outputFile
        "Computer Model: $($computerSystem.Model)" | Add-Content -Path $outputFile
        "Serial Number: $($biosInfo.SerialNumber)" | Add-Content -Path $outputFile
        "Processor: $global:ProcessorInfo" | Add-Content -Path $outputFile
        
        # Determine slot count explicitly
        $memModules = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue
        $memSlotCount = 0
        if ($memorySlots) {
            if ($memorySlots -is [array]) {
                foreach ($slot in $memorySlots) {
                    $memSlotCount += $slot.MemoryDevices
                }
            } else {
                $memSlotCount = $memorySlots.MemoryDevices
            }
        }
        
        $memUsedCount = if ($memModules -is [array]) { $memModules.Count } else { if ($memModules) { 1 } else { 0 } }
        
        # Ensure we have valid numbers
        if ($memUsedCount -lt 0) { $memUsedCount = 0 }
        if ($memSlotCount -lt $memUsedCount) { $memSlotCount = $memUsedCount }
        
        # Force output to include used count even if it's 1
        "Total RAM: $global:RamInfo ($($memUsedCount.ToString()) of $($memSlotCount.ToString()) slots used)" | Add-Content -Path $outputFile
        "Storage: $global:StorageInfo" | Add-Content -Path $outputFile
        "Graphics: $global:GraphicsInfo" | Add-Content -Path $outputFile
        "Network Adapters: $($networkAdapters.Count)" | Add-Content -Path $outputFile
        "Bluetooth: $(if ($global:bluetoothFound) { "Yes" } else { "No" })" | Add-Content -Path $outputFile
    }
    catch {
        Write-ErrorInfo -Component "Hardware Summary" -ErrorMsg $_.Exception.Message
        
        # Emergency direct detection
        try {
            "`nEmergency Hardware Summary:" | Add-Content -Path $outputFile
            "Computer Name: $env:COMPUTERNAME" | Add-Content -Path $outputFile
            
            # Direct queries for critical components
            $cpuDirect = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($cpuDirect -and $cpuDirect.Name) { 
                "Processor: $($cpuDirect.Name)" | Add-Content -Path $outputFile 
            } else {
                "Processor: Unknown" | Add-Content -Path $outputFile
            }
            
            $gpuDirect = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -notlike "*MS Idd*" -and $_.CurrentHorizontalResolution -gt 0 } | 
                Select-Object -First 1
            if ($gpuDirect -and $gpuDirect.Name) {
                "Graphics: $($gpuDirect.Name)" | Add-Content -Path $outputFile
            } else {
                "Graphics: Unknown" | Add-Content -Path $outputFile
            }
            
            $diskDirect = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue
            if ($diskDirect) {
                "Storage: $($diskDirect.Count) drives" | Add-Content -Path $outputFile
            } else {
                "Storage: Unknown" | Add-Content -Path $outputFile
            }
        } catch {
            # Last resort - do nothing
        }
    }
}
catch {
    "A critical error occurred during script execution: $($_.Exception.Message)" | Add-Content -Path $outputFile
}
finally {
    # Add completion timestamp
    "`n`nReport completed at: $(Get-Date)" | Add-Content -Path $outputFile
    Write-Host "Hardware inventory complete. Report saved to: $outputFile" -ForegroundColor Green
}