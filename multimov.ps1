[CmdletBinding()]
param(
    [String]$fol
)

Add-Type -AssemblyName System.Windows.Forms

#region func
<#
#>

<#
https://hinchley.net/articles/creating-a-key-logger-via-a-global-system-hook-using-powershell
https://stackoverflow.com/questions/54236696/how-to-capture-global-keystrokes-with-powershell
#>
    $code = '
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
          IntPtr moduleHandle = GetModuleHandle(System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName);
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
    '

    Add-Type -TypeDefinition $code -ReferencedAssemblies System.Windows.Forms -PassThru | Out-Null

#while ($true) {
#    $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
#    if ($key -eq "X") {
#        Write-Host "Do something now."
#    }
#    start-sleep -Milliseconds 300
#}

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
    if ( $null -eq $h )
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
    if ( $null -eq $h )
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

function Get-RandomSort{
    param(  
    [Parameter(
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String[]]$arr,
    [boolean]$random
    ) 
    process {
        $cpyarr = [System.Collections.ArrayList]::new()
        $retarr = [System.Collections.ArrayList]::new()

        $arr|ForEach-Object{
            [void]$cpyarr.add($_)
        }

        if ($random){
            do {
                $i =  $cpyarr.count - 1
                if ($i -gt 0){
                    $rnd = get-random -Minimum 0 -Maximum $i
                } else {
                    $rnd = 0
                }
                $retarr += $cpyarr[$rnd]
                $cpyarr.RemoveAt($rnd)
                $i--
            } while ($i -ge 0)
        } else {
            $retarr=$cpyarr
        }
        return $retarr
    }
}

function killprocess($processpath){
    stop-process -name (Get-ChildItem $processpath).BaseName -ErrorAction SilentlyContinue
}

function get-VideoResolution($source){
## https://automationadmin.com/2017/12/ps-get-video-file-resolution-from-folder/
## sample:
## 
##    $res = get-VideoResolution "C:\Users\foo\Videos\bar.mp4"
##    "width,height={0},{1}" -f $res.A_フレーム幅,$res.A_フレーム高
##

    if (Test-Path -LiteralPath $Source -PathType Container){
        $Sourcefolder=$Source
        $Sourcefile=''
    } else {
        $Sourcefolder=Split-Path $Source -Parent
        $Sourcefile=Split-Path $Source -Leaf
    }

    $Objshell = New-Object -Comobject Shell.Application 

    $Filelist = @() 
    $Attrlist = @{} 
    #$Details = ( "Frame Height", "Frame Width", "Frame Rate" ) # depends on OS language...
    $Details = ( "フレーム高", "フレーム幅", "フレーム率" ) 
   
    $Objfolder = $Objshell.Namespace($Sourcefolder) 
    For ($Attr = 0 ; $Attr -Le 500; $Attr++) 
    { 
        $Attrname = $Objfolder.Getdetailsof($Objfolder.Items, $Attr) 
        If ( $Attrname -And ( -Not $Attrlist.Contains($Attrname) )) 
        {  
            $Attrlist.Add( $Attrname, $Attr )  
        } 
    } 
   
    Foreach ($File In $Objfolder.Items()) 
    {
        if (($Sourcefile -eq '') -or ($Sourcefile -eq $File.Name)){

            Foreach ( $Attr In $Details) 
            { 
                $Attrvalue = $Objfolder.Getdetailsof($File, $Attrlist[$Attr]) 
                If ( $Attrvalue )  
                {  
                    Add-Member -Inputobject $File -Membertype Noteproperty -Name $("A_" + $Attr) -Value $Attrvalue 
                }  
            } 
            $Filelist += $File 
        }
    } 
    $Filelist
}

function Get-CyclicSubArray {
    param (
        [Parameter(Mandatory=$true)]
        [array]$InputArray,
        
        [Parameter(Mandatory=$true)]
        [int]$StartIndex,
        
        [Parameter(Mandatory=$true)]
        [int]$Count
    )
    <#
    指定配列から開始位置と取得する要素数を指定し、末尾を超過した場合は配列の先頭から続きを取得する関数
    #>

    # 配列の長さを取得
    $ArrayLength = $InputArray.Length

    # 結果を格納する配列を初期化
    $ResultArray = @()

    # 指定された数の要素を取得
    for ($i = 0; $i -lt $Count; $i++) {
        # 現在のインデックスを計算（循環的に）
        $CurrentIndex = ($StartIndex + $i) % $ArrayLength
        
        # 要素を結果配列に追加
        $ResultArray += $InputArray[$CurrentIndex]
    }

    return $ResultArray
}
#endregion

#region main
<# multimov #>

enum movieorienttype {
    guess = 0
    yoko  = 1
    tate  = 2
}

if ($fol -eq ""){
    Write-Warning "folder not specified"
    exit
}

if (!(Test-Path $fol)){
    Write-Warning "Not:found: $fol"
    exit
} 

$filter="*.*"
$loopflg = $true
$movlist_pos=0

$mov_maxcount = 3
$movieorient = [movieorienttype]::guess #0=guess 1=yoko 2=tate
$horizontal_mov_mincount = 6 #horizontal=yoko
$vertical_mov_mincount = 1   #vertical=tate
$h_tilecount = 0
$v_tilecount = 0
$winborder = 8 # depends on environment/settings?
$keys_exit=@('Q','Space','Escape')
$keys_next=@('N','Emter')
$random=$false
#$random=$true

$movlistall = (,((Get-ChildItem $fol -file -Filter $filter).FullName))|Get-RandomSort -random $random

#guess movie orient from movlistall
if ($movieorient -eq [movieorienttype]::guess){
    $h_mov_count = 0
    $v_mov_count = 0
    $movlistall | Select-Object -First $mov_maxcount | ForEach-Object {
        $res = get-VideoResolution $_
        if ([int]$res.A_フレーム幅 -gt [int]$res.A_フレーム高){
            $h_mov_count++
        } else {
            $v_mov_count++
        }
        #"dbg: {1},{2} h={3} v={4}:{0}" -f $res.name,$res.A_フレーム幅,$res.A_フレーム高, $h_mov_count, $v_mov_count
    }
    if ($h_mov_count -ge $v_mov_count){
        $movieorient = [movieorienttype]::yoko
    } else {
        $movieorient = [movieorienttype]::tate
    }
    write-host ("guessing movie orient ... horizontal,vertical={0},{1} => {2}" -f $h_mov_count,$v_mov_count,[movieorienttype].GetEnumName($movieorient))
}

switch ($movieorient){
    ([movieorienttype]::yoko) {
        $mov_count = [math]::Ceiling($mov_maxcount/$vertical_mov_mincount) * $vertical_mov_mincount
    }
    ([movieorienttype]::tate) {
        $mov_count = [math]::Ceiling($mov_maxcount/$horizontal_mov_mincount) * $horizontal_mov_mincount
    }
}


"dbg: 0 mov_count=$mov_count mov_maxcount=$mov_maxcount movieorient=$movieorient"
#pause

switch ($movieorient){
    ([movieorienttype]::yoko) {
	    $v_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
        if ($v_tilecount -lt $vertical_mov_mincount){
            $v_tilecount = $vertical_mov_mincount
        }

        $h_tilecount = [Math]::Ceiling($mov_count / $v_tilecount)

        if ($h_tilecount -eq 1){
            if ($v_tilecount -gt $mov_maxcount){
                $v_tilecount = $mov_maxcount
                $mov_count = $mov_maxcount
            }
        } else {
            $mov_count = $h_tilecount * $v_tilecount
        }

    }
    ([movieorienttype]::tate) {
	    $h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
        if ($h_tilecount -lt $horizontal_mov_mincount){
            $h_tilecount = $horizontal_mov_mincount
        }

        $v_tilecount = [Math]::Ceiling($mov_count / $h_tilecount)

        if ($v_tilecount -eq 1){
            if ($h_tilecount -gt $mov_maxcount){
                $h_tilecount = $mov_maxcount
                $mov_count = $mov_maxcount
            }
        } else {
            $mov_count = $h_tilecount * $v_tilecount
        }
    }
}


# get display size nad taskbar size
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

# get movie width,height from display size, h/v tilecount
$mov_width = [Math]::Ceiling(($width + ($winborder * 2 * $h_tilecount)) / $h_tilecount)
$mov_height = [Math]::Ceiling(($height + ($winborder * $v_tilecount)) / $v_tilecount)

#$mov = $movlistall | Select-Object -First $mov_maxcount

"dbg:mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height"


while ($loopflg){

    #$mov = $movlistall[$movlist_pos..  ($movlist_pos + $mov_maxcount)]
    $mov = Get-CyclicSubArray $movlistall $movlist_pos $mov_count


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
        '--width=0',
        '--height=0',
        "--video-x=0",
        "--video-y=0",
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

    $mov|ForEach-Object{
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
        $mov|ForEach-Object{
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
    $vlcarr|ForEach-Object{
        $mainwindowhandle = (Get-Process -Id $_.id).MainWindowHandle
        Write-Debug "dbg: $($_.id) $mainwindowhandle"
        if ($mainwindowhandle -ne 0){
            $i++
        }
    }
    Write-Verbose "vlc handle found=$i"
    Start-Sleep -Milliseconds 300
}while ($i -lt $movcount)

# arrange window position and size
[int]$x=$x_start
[int]$y=$y_start
[int]$w=$mov_width
[int]$h=$mov_height
$i=0
$vlcarr|ForEach-Object{
    #Set-Window -Id $_.id -X $x -Y $y -Width $w -Height $h -Passthru -Verbose
    Set-Window -Id $_.id -X $x -Y $y -Width $w -Height $h
    $x += $w - ($winborder * 2)
    $i++
    if ($i -ge $h_tilecount){
        $x = $x_start
        $y += $h - $winborder
        $i -= $h_tilecount
    }
}


    #####
    $playflg=$true

    while ($playflg) {
        $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
        "[$key]"
        if ($key -in $keys_exit) {
            Write-Host "quitting.."
            killprocess $vlcPath
            $playflg=$false
            $loopflg=$false
        }
        if ($key -in $keys_next) {
            Write-Host "next.."
            $playflg=$false
            $movlist_pos += $mov_count
        }

        Start-Sleep -Milliseconds 300
    }

}
