Function New-GroupMemberReport {

        
        <#
        requires -module EnhancedHTML2
        .SYNOPSIS
        Generates an HTML-based report for one or more AD groups.
        All groups will be in a table that is hidden and can be expanded individually
        
        .PARAMETER Group
        The group(s) to report members for.
        
        .PARAMETER Path
        The path of the folder where the files should be written.

        .EXAMPLE
        New-GroupMemberReport -Group "Domain Admins","SSLVPN-Users","RDGatewayUsers" -Path c:\Reports

        .EXAMPLE
        "Domain Admins","SSLVPN-Users","RDGatewayUsers" | New-GroupMemberReport -Path c:\Reports
        
        
        #>
        [CmdletBinding()]
        param(
            [Parameter(
                Mandatory=$True,
                ValueFromPipeline=$True,
                ValueFromPipelineByPropertyName=$True)]
            [string[]]$Group,

            [Parameter(Mandatory=$True)]
            [string]$Path
        )

        BEGIN
        {
            if(($rsat = Get-WindowsCapability -Name rsat.active* -Online).state -ne 'Installed')
            {
                Add-WindowsCapability -Name $rsat.name -Online
            }
        
            if(Get-Module -Name EnhancedHTML2 -ErrorAction SilentlyContinue)
            {
                Remove-Module EnhancedHTML2
            }
            else
            {
                $url = "https://www.powershellgallery.com/api/v2/package/EnhancedHTML2/2.0"
                $output = Join-Path $env:TEMP -ChildPath "enhancedhtml2.1.0.1.zip"
            
                $wc = New-Object System.Net.WebClient

                Try
                {
                    $wc.DownloadFile($url, $output)
                }
                catch
                {}

                $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

                if($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
                {
                    $extractpath = Join-Path $env:ProgramFiles -ChildPath 'WindowsPowerShell\Modules\EnhancedHTML2\2.1.0.1'
                }
                else
                {
                    $extractpath = Join-Path $env:USERPROFILE -ChildPath 'Documents\WindowsPowerShell\Modules\EnhancedHTML2\2.1.0.1'
                }

                try
                {
                    Expand-Archive -Path $output -DestinationPath $extractpath
                    Start-Sleep -Seconds 2
                    Remove-Item $output
                }
                catch
                {}
            
            }
            $modulefile = Join-Path $extractpath -ChildPath "EnhancedHTML2.psm1"
            Import-Module $modulefile


        $style = @"
            body {
                color:#333333;
                font-family:Calibri,Tahoma;
                font-size: 12pt;
            }

            h1 {
                color:#003BD8;
                text-align:center;
                font-size: 20pt;
            }

            h2 {
                border-top:1px solid #666666;
                font-size: 16pt;
            }

            th {
                font-weight:bold;
                color:#eeeeee;
                background-color:#333333;
                cursor:pointer;
            }

            .odd  { background-color:#ffffff; }

            .even { background-color:#dddddd; }

            .paginate_enabled_next, .paginate_enabled_previous {
                cursor:pointer; 
                border:1px solid #222222; 
                background-color:#dddddd; 
                padding:2px; 
                margin:4px;
                border-radius:2px;
            }

            .paginate_disabled_previous, .paginate_disabled_next {
                color:#666666; 
                cursor:pointer;
                background-color:#dddddd; 
                padding:2px; 
                margin:4px;
                border-radius:2px;
            }

            .dataTables_info { margin-bottom:4px; }

            .sectionheader { cursor:pointer; }

            .sectionheader:hover { color:red; }

            .grid { width:100% }

            .red {
                color:red;
                font-weight:bold;
            }
"@

            function Get-GroupMembers {
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory=$True)][string]$Group
                )
                    $grp = Get-ADGroup $Group
            
                    $members = Get-ADGroupMember -Identity $grp.distinguishedname

                    $FullMembersDetails = foreach($member in $members)
                    {
                        if($member.objectClass -eq 'User')
                        {
                            $user = Get-ADUser $member -Properties Displayname,Name,SamAccountname,EmailAddress,Enabled,created
                            $props = [ordered]@{
                                Displayname     = $user.Displayname
                                Name            = $user.name
                                SamAccountname  = $user.samaccountname
                                "Email Address" = $user.emailaddress
                                "Enabled"       = $user.enabled
                                "Date Created"  = $user.created
                            }
                            New-Object -TypeName PSObject -Property $props
                        }
                        elseif($member.objectClass -eq 'Group')
                        {
                            $membergroup = Get-ADGroup $member -Properties Name,SamAccountname,GroupCategory,GroupScope,Created,DistinguishedName
                            $props = [ordered]@{
                                Displayname     = "$($membergroup.name) (Security Group)"
                                Name            = "$($membergroup.name) (Security Group)"
                                SamAccountname  = "$($membergroup.samaccountname) (Security Group)"
                                "Email Address" = "N/A"
                                "Enabled"       = "N/A"
                                "Date Created"  = $membergroup.created
                            }
                            New-Object -TypeName PSObject -Property $props
                        }
                
                    }

                    $params = @{'As'='Table';
                            'PreContent'="<h2>&rtrif; $($grp.name)</h2>";
                            'EvenRowCssClass'='even';
                            'OddRowCssClass'='odd';
                            'MakeHiddenSection'=$true;
                            'TableCssClass'='grid'}

                    $FullMembersDetails | ConvertTo-EnhancedHTMLFragment @params -Properties @($FullMembersDetails | Get-Member -MemberType Properties | select -ExpandProperty name)
            }

            function Get-GroupComputers {
                [CmdletBinding()]
                param(
                    [Parameter(Mandatory=$True)][string]$Group
                )
                    $grp = Get-ADGroup $Group

                    $members = Get-ADGroupMember -Identity $grp.distinguishedname

                    $FullMembersDetails = foreach($member in $members)
                    {

                        $computer = Get-ADComputer $member -Properties Displayname,Name,IPv4Address,Enabled,OperatingSystem,OperatingSystemServicePack,created,SamAccountname
                        $props = [ordered]@{
                            Displayname        = $computer.SamAccountname
                            Name               = $computer.name
                            "IP Address"       = $computer.IPv4Address
                            "Operating System" = "$($computer.OperatingSystem) $($computer.OperatingSystemServicePack)"
                            "Enabled"          = $computer.enabled
                            "Date Created"     = $computer.created
                        }
                        New-Object -TypeName PSObject -Property $props

                    }

                    $params = @{'As'='Table';
                            'PreContent'="<h2>&rtrif; $($grp.name)</h2>";
                            'EvenRowCssClass'='even';
                            'OddRowCssClass'='odd';
                            'MakeHiddenSection'=$true;
                            'TableCssClass'='grid'}

                    $FullMembersDetails | ConvertTo-EnhancedHTMLFragment @params -Properties @($FullMembersDetails | Get-Member -MemberType Properties | select -ExpandProperty name)
            }

            $reportname = "Group Membership Report - $((get-date).ToShortDateString().replace('/','-')).html"
            $filepath = Join-Path -Path $Path -ChildPath $reportname
        }

        PROCESS
        {

            $fragments = foreach($grp in $group)
            {
                if($grp -eq "rdgatewaycomputers")
                {
                    Get-GroupComputers -Group $grp
                }
                else
                {
                    Get-GroupMembers -Group $grp
                }

            }
            <#
            $params = @{'CssStyleSheet'=$style;
                        'Title'="System Report for $computer";
                        'PreContent'="<h1>System Report for $computer</h1>";
                'HTMLFragments'=@($html_os,$html_cs,$html_dr,$html_pr,$html_sv,$html_na);
                        'jQueryDataTableUri'='C:\html\jquerydatatable.js';
                        'jQueryUri'='C:\html\jquery.js'}
            ConvertTo-EnhancedHTML @params |
            Out-File -FilePath $filepath
            #>
        
            $params = @{'CssStyleSheet'=$style;
                        'Title'="Group Membership Report";
                        'PreContent'="<h1>Group Membership Report created $(Get-Date)</h1>";
                'HTMLFragments'=@($fragments)}
            ConvertTo-EnhancedHTML @params |
            Out-File -FilePath $filepath -Append

        }

        end
        {
            get-item $filepath
        }
    
}
