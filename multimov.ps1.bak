param(
    [String]$fol
)

Add-Type -AssemblyName System.Windows.Forms

#region func

#region hotkey
Add-Type -TypeDefinition '
using System;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KeyLogger {
  public static class Program {
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;

    private static HookProc hookProc = HookCallback;
    private static IntPtr hookId = IntPtr.Zero;
    private static int keyCode = 0;

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    public static int WaitForKey() {
      hookId = SetHook(hookProc);
      Application.Run();
      UnhookWindowsHookEx(hookId);
      return keyCode;
    }

    private static IntPtr SetHook(HookProc hookProc) {
      IntPtr moduleHandle = GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName);
      return SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, moduleHandle, 0);
    }

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
      if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN) {
        keyCode = Marshal.ReadInt32(lParam);
        Application.Exit();
      }
      return CallNextHookEx(hookId, nCode, wParam, lParam);
    }
  }
}
' -ReferencedAssemblies System.Windows.Forms

#while ($true) {
#    $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
#    if ($key -eq "X") {
#        Write-Host "Do something now."
#    }
#}

#endregion

function Get-TaskBarDimensions {
    param (
        [System.Windows.Forms.Screen]$Screen = [System.Windows.Forms.Screen]::PrimaryScreen
    )        

    $device = ($Screen.DeviceName -split '\\')[-1]
    if ($Screen.Primary) { $device += ' (Primary Screen)' }

    if ($Screen.Bounds.Equals($Screen.WorkingArea)) {
        Write-Warning "Taskbar is hidden on device $device or moved to another screen."
        return
    }


    # calculate heights and widths for the possible positions (left, top, right and bottom)
    $ScreenRect  = $Screen.Bounds
    $workingArea = $Screen.WorkingArea
    $left        = [Math]::Abs([Math]::Abs($ScreenRect.Left) - [Math]::Abs($WorkingArea.Left))
    $top         = [Math]::Abs([Math]::Abs($ScreenRect.Top) - [Math]::Abs($workingArea.Top))
    $right       = ($ScreenRect.Width - $left) - $workingArea.Width
    $bottom      = ($ScreenRect.Height - $top) - $workingArea.Height

    if ($bottom -gt 0) {
        # TaskBar is docked to the bottom
        return [PsCustomObject]@{
            X        = $workingArea.Left
            Y        = $workingArea.Bottom
            Width    = $workingArea.Width
            Height   = $bottom
            Position = 'Bottom'
            Device   = $device
        }
    }
    if ($left -gt 0) {
        # TaskBar is docked to the left
        return [PsCustomObject]@{
            X        = $ScreenRect.Left
            Y        = $ScreenRect.Top
            Width    = $left
            Height   = $ScreenRect.Height
            Position = 'Left'
            Device   = $device
        }
    }
    if ($top -gt 0) {
        # TaskBar is docked to the top
        return [PsCustomObject]@{
            X        = $workingArea.Left
            Y        = $ScreenRect.Top
            Width    = $workingArea.Width
            Height   = $top
            Position = 'Top'
            Device   = $device
        }
    }
    if ($right -gt 0) {
        # TaskBar is docked to the right
        return [PsCustomObject]@{
            X        = $workingArea.Right
            Y        = $ScreenRect.Top
            Width    = $right
            Height   = $ScreenRect.Height
            Position = 'Right'
            Device   = $device
        }
    }
}

Function Set-Window {
<#
.SYNOPSIS
Retrieve/Set the window size and coordinates of a process window.

.DESCRIPTION
Retrieve/Set the size (height,width) and coordinates (x,y) 
of a process window.

.PARAMETER ProcessName
Name of the process to determine the window characteristics. 
(All processes if omitted).

.PARAMETER Id
Id of the process to determine the window characteristics. 

.PARAMETER X
Set the position of the window in pixels from the left.

.PARAMETER Y
Set the position of the window in pixels from the top.

.PARAMETER Width
Set the width of the window.

.PARAMETER Height
Set the height of the window.

.PARAMETER Passthru
Returns the output object of the window.

.NOTES
Name:   Set-Window
Author: Boe Prox
Version History:
    1.0//Boe Prox - 11/24/2015 - Initial build
    1.1//JosefZ   - 19.05.2018 - Treats more process instances 
                                 of supplied process name properly
    1.2//JosefZ   - 21.02.2019 - Parameter Id

.OUTPUTS
None
System.Management.Automation.PSCustomObject
System.Object

.EXAMPLE
Get-Process powershell | Set-Window -X 20 -Y 40 -Passthru -Verbose
VERBOSE: powershell (Id=11140, Handle=132410)

Id          : 11140
ProcessName : powershell
Size        : 1134,781
TopLeft     : 20,40
BottomRight : 1154,821

Description: Set the coordinates on the window for the process PowerShell.exe

.EXAMPLE
$windowArray = Set-Window -Passthru
WARNING: cmd (1096) is minimized! Coordinates will not be accurate.

    PS C:\>$windowArray | Format-Table -AutoSize

  Id ProcessName    Size     TopLeft       BottomRight  
  -- -----------    ----     -------       -----------  
1096 cmd            199,34   -32000,-32000 -31801,-31966
4088 explorer       1280,50  0,974         1280,1024    
6880 powershell     1280,974 0,0           1280,974     

Description: Get the coordinates of all visible windows and save them into the
             $windowArray variable. Then, display them in a table view.

.EXAMPLE
Set-Window -Id $PID -Passthru | Format-Table
​‌‍
  Id ProcessName Size     TopLeft BottomRight
  -- ----------- ----     ------- -----------
7840 pwsh        1024,638 0,0     1024,638

Description: Display the coordinates of the window for the current 
             PowerShell session in a table view.



#>
[cmdletbinding(DefaultParameterSetName='Name')]
Param (
    [parameter(Mandatory=$False,
        ValueFromPipelineByPropertyName=$True, ParameterSetName='Name')]
    [string]$ProcessName='*',
    [parameter(Mandatory=$True,
        ValueFromPipeline=$False,              ParameterSetName='Id')]
    [int]$Id,
    [int]$X,
    [int]$Y,
    [int]$Width,
    [int]$Height,
    [switch]$Passthru
)
Begin {
    Try { 
        [void][Window]
    } Catch {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Window {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(
            IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public extern static bool MoveWindow( 
            IntPtr handle, int x, int y, int width, int height, bool redraw);

        [DllImport("user32.dll")] 
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(
            IntPtr handle, int state);
        }
        public struct RECT
        {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
        }
"@
    }
}
Process {
    $Rectangle = New-Object RECT
    If ( $PSBoundParameters.ContainsKey('Id') ) {
        $Processes = Get-Process -Id $Id -ErrorAction SilentlyContinue
    } else {
        $Processes = Get-Process -Name "$ProcessName" -ErrorAction SilentlyContinue
    }
    if ( $null -eq $Processes ) {
        If ( $PSBoundParameters['Passthru'] ) {
            Write-Warning 'No process match criteria specified'
        }
    } else {
        $Processes | ForEach-Object {
            $Handle = $_.MainWindowHandle
            Write-Verbose "$($_.ProcessName) `(Id=$($_.Id), Handle=$Handle`)"
            if ( $Handle -eq [System.IntPtr]::Zero ) { return }
            $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
            If (-NOT $PSBoundParameters.ContainsKey('X')) {
                $X = $Rectangle.Left            
            }
            If (-NOT $PSBoundParameters.ContainsKey('Y')) {
                $Y = $Rectangle.Top
            }
            If (-NOT $PSBoundParameters.ContainsKey('Width')) {
                $Width = $Rectangle.Right - $Rectangle.Left
            }
            If (-NOT $PSBoundParameters.ContainsKey('Height')) {
                $Height = $Rectangle.Bottom - $Rectangle.Top
            }
            If ( $Return ) {
                $Return = [Window]::MoveWindow($Handle, $x, $y, $Width, $Height,$True)
            }
            If ( $PSBoundParameters['Passthru'] ) {
                $Rectangle = New-Object RECT
                $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
                If ( $Return ) {
                    $Height      = $Rectangle.Bottom - $Rectangle.Top
                    $Width       = $Rectangle.Right  - $Rectangle.Left
                    $Size        = New-Object System.Management.Automation.Host.Size        -ArgumentList $Width, $Height
                    $TopLeft     = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left , $Rectangle.Top
                    $BottomRight = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
                    If ($Rectangle.Top    -lt 0 -AND 
                        $Rectangle.Bottom -lt 0 -AND
                        $Rectangle.Left   -lt 0 -AND
                        $Rectangle.Right  -lt 0) {
                        Write-Warning "$($_.ProcessName) `($($_.Id)`) is minimized! Coordinates will not be accurate."
                    }
                    $Object = [PSCustomObject]@{
                        Id          = $_.Id
                        ProcessName = $_.ProcessName
                        Size        = $Size
                        TopLeft     = $TopLeft
                        BottomRight = $BottomRight
                    }
                    $Object
                }
            }
        }
    }
}
}

function Get-Hwnd($id, $instance = 0){
    $h = Get-Process | Where-Object { $_.Id -match $id } | ForEach-Object { $_.Id }
    if ( $h -eq $null )
    {
        return 0
    }
    else
    {
        if ( $h -is [System.Array] )
        {
            $h = $h[$instance]
        }
        return $h
    }
}

function Get-Hwnd-title($winTitle, $instance = 0){
    $h = Get-Process | Where-Object { $_.MainWindowTitle -match $winTitle } | ForEach-Object { $_.MainWindowHandle }
    if ( $h -eq $null )
    {
        return 0
    }
    else
    {
        if ( $h -is [System.Array] )
        {

            $h = $h[$instance]
        }
        return $h
    }
}

#region getwindowrect
Try { 
    [Void][Window]
} Catch {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Window {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr handle, int state);
    }
    public struct RECT {
        public int Left;   // x position of upper-left corner
        public int Top;    // y position of upper-left corner
        public int Right;  // x position of lower-right corner
        public int Bottom; // y position of lower-right corner
    }
    public struct RECT2
    {
      public string Name;
      public int X;
      public int Y;
      public int Width;
      public int Height;
    }
"@
}
filter ConvertTo-Rect2 ($name, $rc)
{
  $rc2 = New-Object RECT2

  $rc2.Name = $name
  $rc2.X = $rc.Left
  $rc2.Y = $rc.Top
  $rc2.Width = $rc.Right - $rc.Left
  $rc2.Height = $rc.Bottom - $rc.Top

  $rc2
}

function get-windowrect_title($title){
    $WH  = Get-Hwnd $title
    $R   = New-Object RECT
    [Void][Window]::GetWindowRect($WH,[ref]$R)
    #return $R
    return ConvertTo-Rect2 $title $R
}

function get-windowrect($id){
    $WH  = (Get-Process -Id $id).MainWindowHandle 
    $R   = New-Object RECT
    [Void][Window]::GetWindowRect($WH,[ref]$R)
    #return $R
    return ConvertTo-Rect2 $title $R
}
#endregion

function sort-random{
    param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String[]]$arr
    ) 
    process {
        $cpyarr = [System.Collections.ArrayList]::new()
        $retarr = [System.Collections.ArrayList]::new()

        $arr|%{
            [void]$cpyarr.add($_)
        }

        do {
            $i =  $cpyarr.count - 1
            $rnd = get-random -Minimum 0 -Maximum $i
            $retarr += $cpyarr[$rnd]
            $cpyarr.RemoveAt($rnd)
            $i--
        } while ($i -gt 0)
        return $retarr
    }
}

function killprocess($processpath){
    stop-process -name (gci $processpath).BaseName -ErrorAction SilentlyContinue
}

#endregion

if ($fol -eq ""){
    Write-Warning "folder not specified"
    exit
}
if (!(Test-Path $fol)){
    Write-Warning "folder not found : $fol"
    exit
}

$mov_maxcount = 12

$mov=(,((gci $fol -file).FullName)) | sort-random | select -First $mov_maxcount

$screen_width=[System.Windows.Forms.SystemInformation]::WorkingArea.Width
$screen_height=[System.Windows.Forms.SystemInformation]::WorkingArea.Height
$taskbar = Get-TaskBarDimensions
if ($taskbar.Position -in ('Top','Bottom')){
    $width = $screen_width
    #$height = $screen_height - $taskbar.height
    $height = $screen_height
    $x_start = 0
    if ($taskbar.Position -eq 'Top'){
        $y_start = $taskbar.Height
    } else {
        $y_start = 0
    }
} else {
    $width = $screen_width - $taskbar.width
    $height = $screen_height
    if ($taskbar.Position -eq 'Left'){
        $x_start = $taskbar.Width
    } else {
        $x_start = 0
    }
    $y_start = 0
}

$movcount = $mov.count
$movieorient = 2 #1=yoko 2=tate
$vertical_mov_mincount = 0
$vertical_mov_maxcount = 6
$h_tilecount = 0
$v_tilecount = 0
$winborder = 8 # depends on environment/settings?

if ($movieorient -eq 1){
    	$h_tilecount = [Math]::Floor([Math]::Sqrt($movcount / $movieorient)) # yoko yuusen
} else {
   	if ($movcount -lt $vertical_mov_maxcount){
    	$h_tilecount = $movcount
    } else {
        $h_tilecount = [Math]::Ceiling($movcount / $vertical_mov_maxcount)
    }
}
$v_tilecount = [Math]::Ceiling($movcount / $h_tilecount)
if ($vertical_mov_mincount -gt $v_tilecount){$v_tilecount = $vertical_mov_mincount}

$mov_width = [Math]::Ceiling(($width + ($winborder * 2 * $v_tilecount)) / $v_tilecount)
$mov_height = [Math]::Ceiling(($height + ($winborder * $h_tilecount)) / $h_tilecount)

Write-Debug "tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height"

#####
$vlcPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"
$vlcarg = @(
    '--qt-minimal-view',
    '--no-video-deco',
    '--no-qt-video-autoresize',
    '--qt-name-in-title',
    '--qt-continue=0',
    '--no-qt-updates-notif',
    '--no-qt-recentplay',
    '--loop',
# not working (bug in vlc3?)
#    '--width=0',
#    '--height=0',
#    "--video-x=0",
#    "--video-y=0",
# unusable. can't continue script
#    '--no-embedded-video',
# unusable. can't change position
#    '--intf=d',
#    '--dummy-quiet',
# not working
#    '--mmdevice-volume=0.100000',
#    '--directx-volume=0.100000',
#    '--waveout-volume=0.100000',
#    '--volume=50',
    '' # this blank is used to set movie file path
)
$vlcarr = [System.Collections.ArrayList]::new()

killprocess $vlcPath

$mov|%{
    $vlcarg[$vlcarg.count-1]="""$_"""
    $vlcarr += Start-Process $vlcpath -ArgumentList $vlcarg -PassThru
}

# wait for vlc processes to start
Do{
    try {
        $vlccount=(Get-Process -Name vlc).Count
    } catch {
    }
    write-verbose "vlccount=$vlccount"
    start-sleep -Milliseconds 300
}while ($vlccount -lt $movcount)

# wait for vlc to open files ... vlc title name uses metadata rather than filename if has one..
$chktitle=$false
if ($chktitle){
    $i=0
    do {
        $i=0
        $mov|%{
            $filename = split-path $_ -Leaf
            $hwnd = Get-Hwnd-title "$filename - VLCメディアプレイヤー"
            if ($hwnd -ne 0){
                $i++
            }
        }
        write-host "window found=$i"
        start-sleep -Milliseconds 300
    }while ($i -lt $movcount)
}

# wait for vlc mainwindowhandle
#  vlc title name uses metadata rather than filename if has one.
#  we cannot use movie_title, 1 or more titles can be null.
#  instead, we check mainwindowhandle!=0 to wait for vlc windows are opened
Write-Host "waiting for vlc windows..."
do {
    $i=0
    $vlcarr|%{
        $mainwindowhandle = (Get-Process -Id $_.id).MainWindowHandle
        Write-Debug "dbg: $($_.id) $mainwindowhandle"
        if ($mainwindowhandle -ne 0){
            $i++
        }
    }
    Write-Verbose "vlc handle found=$i"
    Start-Sleep -Milliseconds 300
}while ($i -lt $movcount)

#start-sleep 3

# arrange window position and size
[int]$x=$x_start
[int]$y=$y_start
[int]$w=$mov_width
[int]$h=$mov_height
$i=0
$vlcarr|%{
    #Set-Window -Id $_.id -X $x -Y $y -Width $w -Height $h -Passthru -Verbose
    Set-Window -Id $_.id -X $x -Y $y -Width $w -Height $h
    $x += $w - ($winborder * 2)
    $i++
    if ($i -ge $v_tilecount){
        $x = $x_start
        $y += $h - $winborder
        $i -= $v_tilecount
    }
}

#####
$loopflg=$true

while ($loopflg) {
    $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
    if ($key -eq "Q") {
        Write-Host "quitting.."
        killprocess $vlcPath
        $loopflg=$false
    }
}
