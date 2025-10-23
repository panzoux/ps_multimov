<#
.SYNOPSIS
    複数動画ファイルをVLCでタイル表示するスクリプト

.DESCRIPTION
    指定フォルダ内の動画ファイルを取得し、VLCメディアプレイヤーでタイル状に並べて再生します。
    動画の向き（横/縦）を自動判定し、ウィンドウサイズ・位置を自動調整します。
    キー操作で終了(q)や次の動画セットへの切り替え(n)が可能です。

.PARAMETER fol
    動画ファイルが格納されているフォルダのパス

.PARAMETER filter
    動画ファイルのフィルター（例: *.mp4）

.EXAMPLE
    .\multimov.ps1 -fol "C:\Videos" -filter "*.mp4"

.NOTES
    Author: panzoux
    Date: 2025-10-04
#>

    [CmdletBinding()]
param(
    #
    [Parameter(Mandatory=$false)]
    [String]$fol,
    #
    [Parameter(Mandatory=$false)]
    [String]$filter = '*.*',
    #
    [ValidateSet('guess','yoko','tate')]
    [string]$movieorient,
    #
    [Alias('MovMaxCount')]
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,16)]
    [int]$mov_maxcount = 12,
    #
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,8)]
    [int]$HorizontalMinCount = 1,
    #
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,8)]
    [int]$VerticalMinCount   = 1,
    #
    [Parameter(Mandatory=$false)]
    [switch]$Random,
    #
    [Parameter(Mandatory=$false)]
    [switch]$NoLoop
)

enum movieorienttype {
    guess = 0
    yoko  = 1
    tate  = 2
}

Add-Type -AssemblyName System.Windows.Forms

#region func
<##>

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

<#
while ($true) {
    $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
    if ($key -eq "X") {
        Write-Host "Do something now."
    }
    start-sleep -Milliseconds 300
}
#>

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

<#
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

<#
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
#>

#endregion
#>

function Stop-ProcessByPath {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$ProcessPath
    )
    process {
        # extract executable base name if a path was given, else use as-is
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ProcessPath)
        if ([string]::IsNullOrWhiteSpace($baseName)) { return }

        try {
            Get-Process -Name $baseName -ErrorAction SilentlyContinue |
                Stop-Process -ErrorAction SilentlyContinue
        } catch {
            # best-effort fallback
            Stop-Process -Name $baseName -ErrorAction SilentlyContinue
        }
    }
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

<#
powershell  で get-closestdivisor  を実装
第1パラメータで、約数を取得する数値を指定
第2パラメータで、最近値取得用基準値を指定
第3パラメータで、指定値未満最近値、指定値最近値、指定値以上最近値の3パターンを取得できるように
#>

function Get-ClosestDivisor {
    param (
        [int]$Number,        # 約数を取得する数値
        [int]$Reference,     # 最近値取得用基準値
        [ValidateSet("lt", "nearest", "gt")]
        [string]$Mode        # 取得モード
    )

    write-host ("dbg: getclosestdivisor {0} {1} {2}" -f $Number, $Reference, $Mode)

    # 約数を取得するためのリスト
    $divisors = @()

    # 約数を計算
    for ($i = 1; $i -le $Number; $i++) {
        if ($Number % $i -eq 0) {
            $divisors += $i
        }
    }

    # 基準値に基づいて最近値を取得
    switch ($Mode) {
        "lt" {
            $closest = $divisors | Where-Object { $_ -lt $Reference } | Sort-Object -Descending | Select-Object -First 1
        }
        "nearest" {
            $closest = $divisors | Where-Object { $_ -eq $Reference } | Select-Object -First 1
        }
        "gt" {
            $closest = $divisors | Where-Object { $_ -gt $Reference } | Sort-Object | Select-Object -First 1
        }
    }

    return $closest
}

# 使用例
# Get-ClosestDivisor -Number 12 -Reference 5 -Mode "LessThan"
# Get-ClosestDivisor -Number 12 -Reference 6 -Mode "EqualTo"
# Get-ClosestDivisor -Number 12 -Reference 10 -Mode "GreaterThan"

function Get-RectangleMatrix_1{
    param (
        [float]$ContainerWidth,  # 元の矩形の幅
        [float]$ContainerHeight, # 元の矩形の高さ
        [int]$mov_maxcount, # 最大表示数
        [movieorienttype]$movieorient,  # 0=guess 1=yoko 2=tate
        [int]$vertical_mov_mincount,   #vertical=tate
        [int]$horizontal_mov_mincount, #horizontal=yoko
        [int]$winborder
    )

    switch ($movieorient){
        ([movieorienttype]::yoko) {
            $mov_count = [math]::Ceiling($mov_maxcount/$vertical_mov_mincount) * $vertical_mov_mincount
        }
        ([movieorienttype]::tate) {
            $mov_count = [math]::Ceiling($mov_maxcount/$horizontal_mov_mincount) * $horizontal_mov_mincount
        }
    }

    #write-host "dbg: 0 mov_count=$mov_count mov_maxcount=$mov_maxcount movieorient=$movieorient"
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

    if ($h_tilecount -eq 1){
        if ($v_tilecount -gt $mov_maxcount){
            $v_tilecount = $mov_maxcount
            $mov_count = $mov_maxcount
        }
    } else {
        $mov_count = $h_tilecount * $v_tilecount
    }
    if ($v_tilecount -eq 1){
        if ($h_tilecount -gt $mov_maxcount){
            $h_tilecount = $mov_maxcount
            $mov_count = $mov_maxcount
        }
    } else {
        $mov_count = $h_tilecount * $v_tilecount
    }

    $disp = get-displaysize
    $width = $disp.width
    $height = $disp.height
    $winborder = $disp.winborder

    #write-host "dbg: 0 mov_count=$mov_count mov_maxcount=$mov_maxcount movieorient=$movieorient"

    # get movie width,height from display size, h/v tilecount
    $mov_width = [Math]::Ceiling(($width + ($winborder * 2 * $h_tilecount)) / $h_tilecount)
    $mov_height = [Math]::Ceiling(($height + ($winborder * $v_tilecount)) / $v_tilecount)

    "dbg:mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height"

    return [PSCustomObject]@{
        Width            = $mov_width
        Height           = $mov_height
        Columns          = $h_tilecount
        Rows             = $v_tilecount
        Mov_Count        = $mov_count
    }

}

function Get-RectangleMatrix_2 {
    param (
        [float]$ContainerWidth,  # 元の矩形の幅
        [float]$ContainerHeight, # 元の矩形の高さ
        [int]$mov_maxcount,      # 最大表示数
        [movieorienttype]$movieorient,  # 0=guess 1=yoko 2=tate
        [int]$vertical_mov_mincount,   # 縦の最小表示数
        [int]$horizontal_mov_mincount, # 横の最小表示数
        [int]$winborder           # ウィンドウの境界
    )

    # ディスプレイサイズを取得
    $disp = Get-DisplaySize
    $width = $disp.width
    $height = $disp.height
    $winborder = $disp.winborder
    $display_aspect_ratio = $width / $height

    # 表示するタイル数、縦横のタイル数を計算
    switch ($movieorient) {
        ([movieorienttype]::yoko) {
            $mov_count = [math]::Ceiling($mov_maxcount / $vertical_mov_mincount) * $vertical_mov_mincount
            #$h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
            #$h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / $AverageHorizontalRatio * $display_aspect_ratio))
            $h_tilecount = [Math]::Round([Math]::Sqrt($mov_count / $AverageHorizontalRatio * $display_aspect_ratio))
            if ($h_tilecount -lt $horizontal_mov_mincount) {
                $h_tilecount = $horizontal_mov_mincount
            }
            $v_tilecount = [Math]::Ceiling($mov_count / $h_tilecount)
            if ($v_tilecount -lt $vertical_mov_mincount) {
                $v_tilecount = $vertical_mov_mincount
                $h_tilecount = [Math]::Ceiling($mov_count / $v_tilecount)
            }
        }
        ([movieorienttype]::tate) {
            $mov_count = [math]::Ceiling($mov_maxcount / $horizontal_mov_mincount) * $horizontal_mov_mincount
            #$h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
            #$h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / $AverageVerticalRatio * $display_aspect_ratio))
            $h_tilecount = [Math]::Round([Math]::Sqrt($mov_count / $AverageVerticalRatio * $display_aspect_ratio))
            if ($h_tilecount -lt $horizontal_mov_mincount) {
                $h_tilecount = $horizontal_mov_mincount
            }
            $v_tilecount = [Math]::Ceiling($mov_count / $h_tilecount)
            if ($v_tilecount -lt $vertical_mov_mincount) {
                $v_tilecount = $vertical_mov_mincount
                $h_tilecount = [Math]::Ceiling($mov_count / $v_tilecount)
            }
        }
    }

    # タイル数の調整
    if ($h_tilecount -eq 1 -and $v_tilecount -gt $mov_maxcount) {
        $v_tilecount = $mov_maxcount
        $mov_count = $mov_maxcount
    } elseif ($v_tilecount -eq 1 -and $h_tilecount -gt $mov_maxcount) {
        $h_tilecount = $mov_maxcount
        $mov_count = $mov_maxcount
    } else {
        $mov_count = $h_tilecount * $v_tilecount
        if ($mov_maxcount -lt $mov_count){$mov_count = $mov_maxcount} #行*列>全動画数の場合、全動画数を優先(同じ動画の表示を抑制)
    }

    # 動画の幅と高さを計算
    $mov_width = [Math]::Ceiling(($width + ($winborder * 2 * $h_tilecount)) / $h_tilecount)
    $mov_height = [Math]::Ceiling(($height + ($winborder * $v_tilecount)) / $v_tilecount)

    # デバッグ情報
    Write-Verbose "dbg: movieorient=$movieorient mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height"

    # 結果を返す
    return [PSCustomObject]@{
        Width     = $mov_width
        Height    = $mov_height
        Columns   = $h_tilecount
        Rows      = $v_tilecount
        Mov_Count = $mov_count
    }
}

function get-displaysize
{
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

    $winborder = 8 # depends on environment/settings?

    return [PSCustomObject]@{
        Width            = $width
        Height           = $height
        x_start          = $x_start
        y_start          = $y_start
        winborder        = $winborder
    }
}

function get-intFromString
{
    param (
        [string]$inputString
    )

    $result = 0 # 初期値

    # 空文字列の場合は初期値を返す
    if ([string]::IsNullOrWhiteSpace($inputString)) { return $result }

    # 千区切りカンマ／空白を除去してから数字のみ抽出
    $s = $inputString -replace '[,\s]', ''
    if ($s -match '\d+') {
        $num = $matches[0]
        $val = 0
        if ([int]::TryParse($num, [ref]$val)) { return $val }
    }

    return $result
}

function get-movie-orient{
    # guess movie orient and compute average aspect ratios for horizontal/vertical groups
    param (
        $movie_list,
        $movie_orient_type,
        [ref]$AverageHorizontal = $null,
        [ref]$AverageVertical   = $null,
        [ref]$SumAreaHorizontal = $null,
        [ref]$SumAreaVertical   = $null
    )

    # guard: sanitize inputs
    if (-not $movie_list) {
        # nothing to analyze, set defaults and return requested type or guess
        $avgH = 16.0/9.0
        $avgV = 9.0/16.0
        $sumAreaH = 1
        $sumAreaV = 1
        $movie_orient_type_default = [movieorienttype]::yoko
        if ($AverageHorizontal -is [ref]) { $AverageHorizontal.Value = $avgH }
        if ($AverageVertical   -is [ref]) { $AverageVertical.Value   = $avgV }
        if ($SumAreaHorizontal -is [ref]) { $SumAreaHorizontal.Value = $sumAreaH }
        if ($SumAreaVertical   -is [ref]) { $SumAreaVertical.Value   = $sumAreaV }
        $global:AverageHorizontalRatio = $avgH
        $global:AverageVerticalRatio   = $avgV
        $global:SumAreaHorizontal = $sumAreaH
        $global:SumAreaVertical   = $sumAreaV
        if ($movie_orient_type -eq [movieorienttype]::guess){
            Write-Host ("no movies to analyze, defaulting movie orient to {0}  avgRatioH={1} avgRatioV={2}" -f $movie_orient_type_default.ToString(), $avgH, $avgV)
            $movie_orient_type = $movie_orient_type_default
        } else {
            Write-Verbose ("no movies to analyze, using specified movie orient {0}  avgRatioH={1} avgRatioV={2}" -f $movie_orient_type.ToString(), $avgH, $avgV)
        }
        return ([movieorienttype]$movie_orient_type)
    }

    $h_mov_count = 0
    $v_mov_count = 0
    $sumRatioH = 0.0
    $sumRatioV = 0.0
    $sumAreaH = 0
    $sumAreaV = 0

    foreach ($m in $movie_list) {
        $res = Get-VideoResolution $m | Select-Object -First 1
        if (-not $res) { continue }

        $w = get-intFromString $res.'A_フレーム幅'
        $h = get-intFromString $res.'A_フレーム高'

        if ($h -eq 0) { continue }

        #$ratio = [double]$w / [double]$h
        $ratio = [int]$w / [int]$h

        if ($w -gt $h) {
            $h_mov_count++
            $sumRatioH += $ratio
            #$sumAreaH += $w * $h
            $sumAreaH += 100 * (100 * $h / $w) # 横長動画の面積は横幅基準で正規化して加算
        } else {
            $v_mov_count++
            $sumRatioV += $ratio
            #$sumAreaV += $w * $h
            $sumAreaV += (100 * $w / $h) * 100 # 縦長動画の面積は縦幅基準で正規化して加算
        }
        #Write-Verbose ("movie: {0}  res={1}x{2} ratio={3}  h_mov_count={4} v_mov_count={5}" -f $m, $w, $h, [math]::Round($ratio,4), $h_mov_count, $v_mov_count)
    }

    $avgH = if ($h_mov_count -gt 0) { [math]::Round($sumRatioH / $h_mov_count, 4) } else { 1 }
    $avgV = if ($v_mov_count -gt 0) { [math]::Round($sumRatioV / $v_mov_count, 4) } else { 1 }

    # return averages via byref parameters if provided
    if ($AverageHorizontal -is [ref]) { $AverageHorizontal.Value = $avgH }
    if ($AverageVertical   -is [ref]) { $AverageVertical.Value   = $avgV }
    if ($SumAreaHorizontal -is [ref]) { $SumAreaHorizontal.Value = $sumAreaH }
    if ($SumAreaVertical   -is [ref]) { $SumAreaVertical.Value   = $sumAreaV }

    $global:AverageHorizontalRatio = $avgH
    $global:AverageVerticalRatio   = $avgV
    $global:SumAreaHorizontal = $sumAreaH
    $global:SumAreaVertical   = $sumAreaV

    # decide movie orient type if set to guess
    if ($movie_orient_type -eq [movieorienttype]::guess){
        if ($h_mov_count -ge $v_mov_count){
            $movie_orient_type = [movieorienttype]::yoko
        } else {
            $movie_orient_type = [movieorienttype]::tate
        }
        Write-Host ("guessing movie orient ... horizontal,vertical={0},{1} => {2}  avgRatioH={3} avgRatioV={4}" -f $h_mov_count,$v_mov_count,$movie_orient_type.ToString(), $avgH, $avgV)
    } else {
        Write-Verbose ("movie_orient_type specified: {0}  avgRatioH={1} avgRatioV={2}" -f $movie_orient_type.ToString(), $avgH, $avgV)
    }

    return $movie_orient_type
}

#endregion

#region main
<# multimov #>

##### normalize param
# Normalize movie orient param (string) to enum variable used by script
if ($PSBoundParameters.ContainsKey('movieorient') -and -not [string]::IsNullOrWhiteSpace($movieorient)) {
    try {
        $movie_orient_type_param = [movieorienttype]$movieorient
    } catch {
        Write-Verbose "Invalid movieorient '$movieorient', defaulting to 'guess'"
        $movie_orient_type_param = [movieorienttype]::guess
    }
} else {
    $movie_orient_type_param = [movieorienttype]::guess
}

$horizontal_mov_mincount = $HorizontalMinCount
$vertical_mov_mincount = $VerticalMinCount


##### check param
Write-Verbose "dbg:`n param fol=$fol`n filter=$filter`n mov_maxcount=$mov_maxcount`n movie_orient_type_param=$movie_orient_type_param`n vertical_mov_mincount=$vertical_mov_mincount`n horizontal_mov_mincount=$horizontal_mov_mincount` random=$random`n noloop=$noloop"

if ($fol -eq ""){
    Write-Warning "folder not specified"
    exit
}

if (!(Test-Path $fol)){
    Write-Warning "Not:found: $fol"
    exit
} 

if ($filter -eq "") {$filter="*.*"}


##### settings
$keys_exit=@('Q','Space','Escape')
$keys_next=@('N','Enter')


##### vlc settings
$vlcPath = "C:\Program Files\VideoLAN\VLC\vlc.exe"
$vlcarg = @(
    '--qt-minimal-view',
    '--no-video-deco',
    '--no-qt-video-autoresize',
    '--qt-name-in-title',
    '--qt-continue=0',
    '--no-qt-updates-notif',
    '--no-qt-recentplay',
#    '--loop',
# not working (bug in vlc3?)
    '--width=0',
    '--height=0',
    "--video-x=0",
    "--video-y=0"
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
#    '' # this blank is used to set movie file path
)

if ($NoLoop -eq $false){
    $vlcarg+='--loop'
}

$vlcarg+='' # this blank is used to set movie file path

Write-Verbose "dbg: vlcarg=$vlcarg"

#####
$continueLoop = $true
$movlist_pos = 0
$h_tilecount = 0
$v_tilecount = 0
$avgH = 0
$avgV = 0
$sumAreaH = 0
$sumAreaV = 0

# only enumerate common video extensions (unless a specific filter is provided)
$videoExtensions = @('*.mp4','*.mkv','*.avi','*.mov','*.wmv','*.flv','*.webm','*.m4v','*.mpeg','*.mpg','*.ts','*.m2ts','*.3gp')

if ($PSBoundParameters.ContainsKey('filter') -and -not [string]::IsNullOrWhiteSpace($filter) -and $filter -ne '*.*') {
    $patterns = @($filter)
} else {
    $patterns = $videoExtensions
}

$movlistall = forEach ($p in $patterns) {
    Get-ChildItem -Path $fol -File -Filter $p -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
}
$movlistall = $movlistall | Sort-Object -Unique

if ($movlistall.count -eq 0){
    Write-Warning "No movie files found in $fol"
    exit
}

if ($random){
    $movlistall = $movlistall|Get-Random -Count $movlistall.count
}

if ($movlistall.count -lt $mov_maxcount){$mov_count = $movlistall.count} else {$mov_count = $mov_maxcount}

while ($continueLoop){

    $mov = Get-CyclicSubArray $movlistall $movlist_pos $mov_count

    #$movieorient = get-movie-orient -movie_list @($mov) -movie_orient_type $movie_orient_type_param
    $movieorient = get-movie-orient -movie_list @($mov) -movie_orient_type $movie_orient_type_param -AverageHorizontal ([ref]$avgH) -AverageVertical ([ref]$avgV) -SumAreaHorizontal ([ref]$sumAreaH) -SumAreaVertical ([ref]$sumAreaV)
 
    $disp = get-displaysize
    $x_start = $disp.x_start
    $y_start = $disp.y_start
    $winborder = $disp.winborder

    #$rmatrix = Get-RectangleMatrix_1 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_maxcount -movieorient $movieorient -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount
    #$rmatrix = Get-RectangleMatrix_2 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_count -movieorient $movieorient -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount

    # select the matrix with tilecount closer to mov_count
    $rmatrix_v = Get-RectangleMatrix_2 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_count -movieorient ([movieorienttype]::tate).value__ -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount
    $rmatrix_h = Get-RectangleMatrix_2 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_count -movieorient ([movieorienttype]::yoko).value__ -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount
    $diff_movcount_v = [math]::Abs($rmatrix_v.Columns*$rmatrix_v.Rows - $mov_count)
    $diff_movcount_h = [math]::Abs($rmatrix_h.Columns*$rmatrix_h.Rows - $mov_count)

    if (($sumAreaH -eq $sumAreaV) -or ($sumAreaH -lt 1) -or ($sumAreaV -lt 1)){
        # if both sum areas are equal (or invalid), select by movcount diff
        if ($diff_movcount_v -eq $diff_movcount_h){
            # if both diffs are equal, select the one with larger area      
            $area_v = $rmatrix_v.Width * $rmatrix_v.Height
            $area_h = $rmatrix_h.Width * $rmatrix_h.Height
            if ($area_v -eq $area_h){
                # if both areas are equal, select the one with aspect ratio closer to average
                [double]$avgRatio = 0.0
                if ($movieorient -eq [movieorienttype]::yoko){
                    $avgRatio = $avgH
                } else {
                    $avgRatio = $avgV
                }
                [double]$diff_v = [math]::Abs( ($rmatrix_v.Width / $rmatrix_v.Height) - $avgRatio )
                [double]$diff_h = [math]::Abs( ($rmatrix_h.Width / $rmatrix_h.Height) - $avgRatio )
                if ($diff_v -le $diff_h){
                    $rmatrix = $rmatrix_v
                    write-verbose ("dbg: selected v-oriented matrix by aspect ratio: diff_v=$diff_v diff_h=$diff_h")
                } else {
                    $rmatrix = $rmatrix_h
                    write-verbose ("dbg: selected h-oriented matrix by aspect ratio: diff_v=$diff_v diff_h=$diff_h")
                }
            } elseif ($area_v -gt $area_h){
                $rmatrix = $rmatrix_v
                write-verbose ("dbg: selected v-oriented matrix by area: area_v=$area_v area_h=$area_h")
            } else {
                $rmatrix = $rmatrix_h
                write-verbose ("dbg: selected h-oriented matrix by area: area_v=$area_v area_h=$area_h")
            }
        } elseif ($diff_movcount_v -lt $diff_movcount_h){
            $rmatrix = $rmatrix_v
            write-verbose ("dbg: selected v-oriented matrix: diff_movcount_v=$diff_movcount_v diff_movcount_h=$diff_movcount_h")
        } else {
            $rmatrix = $rmatrix_h
            write-verbose ("dbg: selected h-oriented matrix: diff_movcount_v=$diff_movcount_v diff_movcount_h=$diff_movcount_h")
        }
    } elseif ($sumAreaH -gt $sumAreaV){
        # if horizontal sum area is larger, prefer h-oriented matrix
        $rmatrix = $rmatrix_h
        write-verbose ("dbg: selected h-oriented matrix by sum area: sumAreaH=$sumAreaH sumAreaV=$sumAreaV")
    } else {
        # if vertical sum area is larger, prefer v-oriented matrix
        $rmatrix = $rmatrix_v
        write-verbose ("dbg: selected v-oriented matrix by sum area: sumAreaH=$sumAreaH sumAreaV=$sumAreaV")
    }
    write-verbose ("dbg: computed matrix v/h={0}/{1} for mov_count={2}" -f $rmatrix.Rows, $rmatrix.Columns, $mov_count)

    # get movie window size and tilecount
    $mov_width = $rmatrix.Width
    $mov_height = $rmatrix.Height
    $h_tilecount = $rmatrix.Columns
    $v_tilecount = $rmatrix.Rows
    $mov_count = $rmatrix.Mov_Count
    write-verbose "dbg: mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height xyoffset=$($disp.x_start),$($disp.y_start)"

    $vlcarr = [System.Collections.ArrayList]::new()

    Stop-ProcessByPath $vlcPath

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
    }while ($vlccount -lt $mov_count)


    # wait for vlc mainwindowhandle
    #  VLC window titles may use metadata instead of the filename.
    #  Movie titles can be null, so we cannot rely on them.
    #  Instead, we check that MainWindowHandle is not zero to ensure VLC windows have opened.
    Write-Host "waiting for vlc windows..."
    $vlcarr|ForEach-Object{
        do {
            $mainwindowhandle = (Get-Process -Id $_.id).MainWindowHandle
            #Write-verbose "dbg: $($_.id) $mainwindowhandle"
            Start-Sleep -Milliseconds 300
        }while ($mainwindowhandle -eq 0)
    }


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
        #"dbg {0} {1},{2} {3},{4}" -f $_.Id,$x,$y,$w,$h
    }


    #####
    $playflg=$true

    while ($playflg) {
        $key = [System.Windows.Forms.Keys][KeyLogger.Program]::WaitForKey()
        "[$key]"
        if ($key -in $keys_exit) {
            Write-Host "quitting.."
            Stop-ProcessByPath $vlcPath
            $playflg=$false
            $continueLoop=$false
        }
        if ($key -in $keys_next) {
            Write-Host "next.."
            $playflg=$false
            $movlist_pos += $mov_count
        }

        Start-Sleep -Milliseconds 300
    }

}
