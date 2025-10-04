<#
.SYNOPSIS
    複数動画ファイルをVLCでタイル表示するスクリプト

.DESCRIPTION
    指定フォルダ内の動画ファイルを取得し、VLCメディアプレイヤーでタイル状に並べて再生します。
    動画の向き（横/縦）を自動判定し、ウィンドウサイズ・位置を自動調整します。
    キー操作で終了や次の動画セットへの切り替えが可能です。

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
    [String]$fol,
    [String]$filter
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

    write-host "dbg: 0 mov_count=$mov_count mov_maxcount=$mov_maxcount movieorient=$movieorient"
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

    # 表示するタイル数、縦横のタイル数を計算
    switch ($movieorient) {
        ([movieorienttype]::yoko) {
            $mov_count = [math]::Ceiling($mov_maxcount / $vertical_mov_mincount) * $vertical_mov_mincount
            $v_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
            if ($v_tilecount -lt $vertical_mov_mincount) {
                $v_tilecount = $vertical_mov_mincount
            }
            $h_tilecount = [Math]::Ceiling($mov_count / $v_tilecount)
        }
        ([movieorienttype]::tate) {
            $mov_count = [math]::Ceiling($mov_maxcount / $horizontal_mov_mincount) * $horizontal_mov_mincount
            $h_tilecount = [Math]::Ceiling([Math]::Sqrt($mov_count / 2))
            if ($h_tilecount -lt $horizontal_mov_mincount) {
                $h_tilecount = $horizontal_mov_mincount
            }
            $v_tilecount = [Math]::Ceiling($mov_count / $h_tilecount)
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

    # ディスプレイサイズを取得
    $disp = Get-DisplaySize
    $width = $disp.width
    $height = $disp.height
    $winborder = $disp.winborder

    # 動画の幅と高さを計算
    $mov_width = [Math]::Ceiling(($width + ($winborder * 2 * $h_tilecount)) / $h_tilecount)
    $mov_height = [Math]::Ceiling(($height + ($winborder * $v_tilecount)) / $v_tilecount)

    # デバッグ情報
    write-host "dbg: mov_maxcount=$mov_maxcount movieorient=$movieorient"
    write-host "dbg: mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height"

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

function get-movie-orient{
    #guess movie orient
    param (
        $movie_list,$movie_orient_type
    )

    if ($movie_orient_type -eq [movieorienttype]::guess){
        $h_mov_count = 0
        $v_mov_count = 0
        #$movie_list | Select-Object -First $mov_maxcount | ForEach-Object {
        $movie_list | ForEach-Object {
            $res = get-VideoResolution $_
            if ([int]$res.A_フレーム幅 -gt [int]$res.A_フレーム高){
                $h_mov_count++
            } else {
                $v_mov_count++
            }
            #"dbg: {1},{2} h={3} v={4}:{0}" -f $res.name,$res.A_フレーム幅,$res.A_フレーム高, $h_mov_count, $v_mov_count
        }
        if ($h_mov_count -ge $v_mov_count){
            $movie_orient_type = [movieorienttype]::yoko
        } else {
            $movie_orient_type = [movieorienttype]::tate
        }
        write-host ("guessing movie orient ... horizontal,vertical={0},{1} => {2}" -f $h_mov_count,$v_mov_count,[movieorienttype].GetEnumName($movie_orient_type))
    }
    return $movie_orient_type
}

#endregion

#region main
<# multimov #>

##### check param
if ($fol -eq ""){
    Write-Warning "folder not specified"
    exit
}

if (!(Test-Path $fol)){
    Write-Warning "Not:found: $fol"
    exit
} 

if ($filter -eq "") {$filter="*.*"}
$loopflg = $true
$movlist_pos=0


##### settings
$mov_maxcount = 12
$movie_orient_type_param = [movieorienttype]::guess #0=guess 1=yoko 2=tate
$horizontal_mov_mincount = 6 #horizontal=yoko
$vertical_mov_mincount = 1   #vertical=tate
$h_tilecount = 0
$v_tilecount = 0
$keys_exit=@('Q','Space','Escape')
$keys_next=@('N','Emter')
$random=$false
$random=$true


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


#####

$movlistall = (,((Get-ChildItem $fol -file -Filter $filter).FullName))|Get-RandomSort -random $random

if ($movlistall.count -lt $mov_maxcount){$mov_count = $movlistall.count} else {$mov_count = $mov_maxcount}

while ($loopflg){

    $mov = Get-CyclicSubArray $movlistall $movlist_pos $mov_count

    $movieorient = get-movie-orient -movie_list @($mov) -movie_orient_type $movie_orient_type_param

    $disp = get-displaysize
    $x_start = $disp.x_start
    $y_start = $disp.y_start
    $winborder = $disp.winborder

    #$rmatrix = Get-RectangleMatrix_1 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_maxcount -movieorient $movieorient -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount
    $rmatrix = Get-RectangleMatrix_2 -ContainerWidth $disp.width -ContainerHeight $disp.height -mov_maxcount $mov_count -movieorient $movieorient -vertical_mov_mincount $vertical_mov_mincount -horizontal_mov_mincount $horizontal_mov_mincount
    $mov_width = $rmatrix.Width
    $mov_height = $rmatrix.Height
    $h_tilecount = $rmatrix.Columns
    $v_tilecount = $rmatrix.Rows
    $mov_count = $rmatrix.Mov_Count
    "dbg:mov_count=$mov_count tilecount v/h=$v_tilecount/$h_tilecount mov_w/h=$mov_width/$mov_height xyoffset=$($disp.x_start),$($disp.y_start)"


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
#  VLC window titles may use metadata instead of the filename.
#  Movie titles can be null, so we cannot rely on them.
#  Instead, we check that MainWindowHandle is not zero to ensure VLC windows have opened.
Write-Host "waiting for vlc windows..."
$vlcarr|ForEach-Object{
    do {
        $mainwindowhandle = (Get-Process -Id $_.id).MainWindowHandle
        Write-verbose "dbg: $($_.id) $mainwindowhandle"
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
