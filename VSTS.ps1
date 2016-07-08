Param(
    [Parameter(Mandatory=$true,  Position=0, ValueFromPipeline=$true)] [string] $buildId,
    [Parameter(Mandatory=$true,  Position=1, ValueFromPipeline=$true)] [string] $collectionUrl,
    [Parameter(Mandatory=$true,  Position=2, ValueFromPipeline=$true)] [string] $teamproject,
    [Parameter(Mandatory=$true,  Position=3, ValueFromPipeline=$true)] [string] $vstsToken,
    [Parameter(Mandatory=$false, Position=4, ValueFromPipeline=$true)] [string] $outputfile = "ReleaseNotes-$buildId.md",
    [Parameter(Mandatory=$false, Position=4, ValueFromPipeline=$true)] [string] $outputjson = "ReleaseNotes-$buildId.json",
    [Parameter(Mandatory=$false, Position=5, ValueFromPipeline=$true)] [string] $outputpath = "."
)

BEGIN {
    
    function Write-Log {
        param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][AllowEmptyString()][string]$Message)
        Write-Verbose -Verbose ("[{0:s}] {1}`r`n" -f (get-date), $Message)
    }
    function Write-VLog {
        param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][AllowEmptyString()][string]$Message)
        Write-Verbose  ("[{0:s}] {1}`r`n" -f (get-date), $Message)
    }
    function Write-Info {
        param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][AllowEmptyString()][string]$Message)
        Write-host ("INFO:    [{0:s}] {1}`r" -f (get-date), $Message) -fore cyan
    }
    function Write-Warn {
        param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][AllowEmptyString()][string]$Message)
        Write-warning ("[{0:s}] {1}`r" -f (get-date), $Message)
    }
    function Write-Success {
        param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][AllowEmptyString()][string]$Message)
        Write-host ("SUCC:    [{0:s}] {1}`r" -f (get-date), $Message) -fore green
    }
    function get-caller {
        [cmdletbinding()]
        param()
        process{ return ((Get-PSCallStack)[1].Command) }
    }


    function Process-Template($template,$builds) 
    {
            
        if ($template.count -gt 0)
        {
            write-Verbose "Processing template"

            # create our work stack and initialise
            $modeStack = new-object  System.Collections.Stack 
            $modeStack.Push([Mode]::BODY)

            # this line is to provide support the the legacy build only template
            # if using a release template it will be reset when processing tags
            $builditem = $builds
            $build = $builditem.build # using two variables for legacy support
            
            #process each line
            For ($index =0; $index -lt $template.Count; $index++)
            {
                $line = $template[$index]
                # get the line change mode if any
                $mode = Get-Mode -line $line

                #debug logging
                #Write-Verbose "$Index[$($mode)]: $line"

                if ($mode -ne [Mode]::BODY)
                {
                    # is there a mode block change
                    if ($modeStack.Peek().Mode -eq $mode)
                    {
                        # this means we have reached the end of a block
                        # need to work out if there are more items to process 
                        # or the end of the block
                        $queue = $modeStack.Peek().BlockQueue;
                        if ($queue.Count -gt 0)
                        {
                            # get the next item and initialise
                            # the variables exposed to the template
                            $item = $queue.Dequeue()
                            # reset the index to process the block
                            $index = $modeStack.Peek().Index
                            switch ($mode)
                            {
                            "WI" {
                                Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting next workitem $($item.id)"
                                $widetail = $item  
                            }
                            "CS" {
                                Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting next changeset/commit $($item.id)"
                                $csdetail = $item 
                            }
                            "BUILD" {
                                Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting next build $($item.build.id)"
                                $builditem = $item
                                $build = $builditem.build # using two variables for legacy support
                            }
                            } #end switch
                        }
                        else
                        {
                            # end of block and no more items, so exit the block
                            $mode = $modeStack.Pop().Mode
                            Write-Verbose "$(Indent-Space -indent $modeStack.Count)Ending block $mode"
                        }
                    } else {
                        # this a new block to add the stack
                        # need to get the items to process and place them in a queue
                        Write-Verbose "$(Indent-Space -indent ($modeStack.Count))Starting block $($mode)"
                    ###    $queue = new-object  System.Collections.Queue  
                        #set the index to jump back to
                        $lastBlockStartIndex = $index       
                        switch ($mode)
                        {
                            "WI" {
                                # store the block and load the first item
                                Create-StackItem -items @($builditem.workItems) -modeStack $modeStack -mode $mode -index $index
                                if ($modeStack.Peek().BlockQueue.Count -gt 0)
                                {
                                    $widetail = $modeStack.Peek().BlockQueue.Dequeue()
                                    Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting first workitem $($widetail.id)"
                                } else {
                                    $widetail = $null
                                }
                            }
                            "CS" {
                                # store the block and load the first item
                                Create-StackItem -items @($builditem.changesets) -modeStack $modeStack -mode $mode -index $index
                                if ($modeStack.Peek().BlockQueue.Count -gt 0)
                                {
                                    $csdetail = $modeStack.Peek().BlockQueue.Dequeue()   
                                    Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting first changeset/commit $($d.id)"
                                
                                } else {
                                    $csdetail = $null
                                }
                                                        
                            }
                            "BUILD" {
                                Create-StackItem -items @($builds) -modeStack $modeStack -mode $mode -index $index
                                if ($modeStack.Peek().BlockQueue.Count -gt 0)
                                {
                                    $builditem = $modeStack.Peek().BlockQueue.Dequeue() 
                                    $build = $builditem.build
                                    Write-Verbose "$(Indent-Space -indent $modeStack.Count)Getting first build $($build.id)"
                                }  else {
                                    $builditem = $null
                                    $build = $null
                                }
                            }
                        }
                    }
                } else
                {
                if ((($modeStack.Peek().mode -eq [Mode]::WI) -and ($widetail -eq $null)) -or 
                    (($modeStack.Peek().mode -eq [Mode]::CS) -and ($csdetail -eq $null)))
                {
                    # there is no data to expand
                    $out += "None"
                } else {
                    # nothing to expand just process the line
                    $out += $line | render
                    $out += "`n"
                    }
                }
            }
            $out
        } else
        {
            write-error "Cannot load template file [$templatefile] or it is empty"
        } 
    }




    function Indent-Space
    {
        param($size =3,$indent =1)
        $upperBound = $size * $indent
        for ($i =1 ; $i -le $upperBound  ; $i++)
        {
            $padding += " "
        } 
        $padding
    }


    function Render() {
        [CmdletBinding()]
        param ( [parameter(ValueFromPipeline = $true)] [string] $str)

        #buggy in V4 seems ok in older and newer
        #$ExecutionContext.InvokeCommand.ExpandString($str)

        "@`"`n$str`n`"@" | iex
    }

    function Get-Template ($templateLocation,$templatefile,$inlinetemplate)
    {
        Write-Verbose "Using template mode [$templateLocation]"

        if ($templateLocation -eq 'File')
        {
            write-Verbose "Loading template file [$templatefile]"
            $template = Get-Content $templatefile
        } else 
        {
            write-Verbose "Using in-line template"
            # it appears as single line we need to split it out
            $template = $inlinetemplate -split "`n"
        }
        
        return $template
    }

    function Get-Mode($line)
    {
        $mode = [Mode]::BODY
        if ($line.Trim() -eq "@@WILOOP@@") {$mode = [Mode]::WI}
        if ($line.Trim() -eq "@@CSLOOP@@") {$mode = [Mode]::CS}
        if ($line.Trim() -eq "@@BUILDLOOP@@") {$mode = [Mode]::BUILD}
        return $mode
    }

    function Create-StackItem($items,$modeStack,$mode,$index)
    {
        # Create a queue of the items
        $queue = new-object  System.Collections.Queue
        # add each item to the queue  
        foreach ($item in @($items))
        {
            $queue.Enqueue($item)
        }
        Write-Verbose "$(Indent-Space -indent ($modeStack.Count +1))$($queue.Count) items"
        # place it on the stack with the blocks mode and start line index
        $modeStack.Push(@{'Mode'= $mode;
                        'BlockQueue'=$queue;
                        'Index' = $index})
    }


if (-not ([System.Management.Automation.PSTypeName]'Mode').Type)
{
    # types to make the switches neater
    Add-Type -TypeDefinition @"
public enum Mode
{
    BODY,
    WI,
    CS,
    BUILD
}
"@
}



[string] $global:collectionUrl = $collectionUrl
[string] $global:teamproject = $teamproject
[string] $global:buildId = $buildId
[string] $global:user = "ReleaseNotesGenerator"
[string] $global:token = $vstsToken

$defaulttemplate = @"
#Release notes for build `$(`$build.id) [`$(`$build.buildNumber)]

| Property           | Detail                |
| ------------------ | -----------           |
| Build Id           | `$(`$build.id)                  |
| Build Number       | `$(`$build.buildNumber)            |
| Build Status       | `$(`$build.status)             |
| Build Result       | `$(`$build.result)             |
| Build Requested By | `$(`$builds.build.requestedBy.displayName)            |
| Build Start Time   | `$("{0:dd/MM/yy HH:mm:ss}" -f [datetime]`(`$(`$builds.build.startTime)))     |
| Report Time        | `$("{0:dd/MM/yy HH:mm:ss}" -f [datetime]`(Get-Date))     |
| Source Branch      | `$(`$build.sourceBranch)     |

###Associated work items  
@@WILOOP@@  
* **`$(`$widetail.fields.'System.WorkItemType') `$(`$widetail.id) ** [Assigned To: `$(`$widetail.fields.'System.AssignedTo')] `$(`$widetail.fields.'System.Title')
@@WILOOP@@  

###Associated change sets/commits  
@@CSLOOP@@  
> **ChangeSet:** `$(`$csdetail.treeId)
>> **Commit ID:** `$(`$csdetail.commitid) 
`$(`$csdetail.comment)    
@@CSLOOP@@  
"@
        

        
        function Invoke-cmd($uri)
        {
            
            write-success "    URI: $uri"
            $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $global:user,$global:token)))
            $ret = Invoke-RestMethod -Uri $uri -Method Get -ContentType "application/json" -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}
            return $ret
        }
        function Get-Build($buildId)
        {
            #write-info "Getting Build Info for Build: $buildId"
            $uri = ('{0}/{1}/_apis/build/builds/{2}/?api-version=2.0' -f $collectionUrl,$teamproject,$buildId )
            $ret = Invoke-cmd($uri)
            if (!($ret.value))
            {
                write-warn "    There are no Builds associated with buildId=$buildId"
            }
            return $ret
        }
        function Get-BuildChangeSets($buildId)
        {
            write-info "Getting Build ChangeSets for Build: $buildId"
            $uri = ('{0}/{1}/_apis/build/builds/{2}/changes?api-version=2.0' -f $collectionUrl,$teamproject,$buildId )
            $ret = Invoke-cmd($uri)
            $csList = @();
            if ($ret.value)
            {
                foreach ($cs in $ret.value)
                {
                    # we can get more detail if the changeset is on VSTS or TFS
                    try {
                        $csList += Invoke-cmd($cs.location)
                    } catch
                    {
                        Write-warning "Unable to get details of changeset/commit as it is not stored in TFS/VSTS"
                        Write-warning "For [$($cs.id)] location [$($cs.location)]"
                        Write-warning "Just using the details we have from the build"
                        $csList += $cs
                    }
                } 
            }
            else
            {
                write-warn "    There are no Builds associated with buildId=$buildId"
            }
            return $csList
        }
        Function Get-BuildWorkItems($buildId)
        {
            write-info "Getting Work items for Build: $buildId"
            $uri = ('{0}/{1}/_apis/build/builds/{2}/workitems?api-version=2.0' -f $collectionUrl,$teamproject,$buildId )
            $ret = Invoke-cmd($uri)
            $wiList = @();
            if ($ret.value)
            {
                foreach ($wi in $ret.value)
                {
                    # we can get more detail if the changeset is on VSTS or TFS
                    try {
                        #[string] $a = Invoke-cmd($wi.url)
                        #write-host "Value of A is: $($wi.location)" 
                        $wiList += Invoke-cmd($wi.url)
                    } catch
                    {
                        Write-warning "Unable to get details of changeset/commit as it is not stored in TFS/VSTS"
                        Write-warning "For [$($wi.id)] location [$($wi.location)]"
                        Write-warning "Just using the details we have from the build"
                        $wiList += $wi
                    }
                } 
                #$wiList
            }
            else
            {
                write-warn "    There are no Work Items associated with buildId=$buildId"
            }
            return $wiList
        }
        
        function Get-BuildDataSet($buildId)
        {
            write-log "Getting build details for buildId [$buildId]"    
            $build = Get-Build($buildId)
            #$build | out-string
            #break;

            #Write-log "Getting associated work items for build [$($buildId)]"
            #Write-log "Getting associated changesets/commits for build [$($buildId)]"

            $build = @{ 'build'=$build;
                        'id'=($buildId);
                        'number'=($build.buildNumber);
                        'status'=($builds.build.status);
                        'result'=($builds.build.result)
                        'requestedBy'=($($builds.build.requestedBy.displayName));
                        'workitems'=(Get-BuildWorkItems($buildId));
                        'changesets'=(Get-BuildChangeSets($buildId))
                        
                    }
            return $build
        }
     
}

PROCESS{

    #Initial Configuration/Cleanup
    $output = join-path $outputpath $outputfile 
    if (Test-Path $output) { remove-item $output }

    #Get API Info
    $builds = Get-BuildDataSet($buildId)
    $inlinetemplate = $defaulttemplate

    #Format Data
    $template = Get-Template -templateLocation $templateLocation -templatefile $templatefile -inlinetemplate $inlinetemplate
    $outputmarkdown = Process-Template -template $template -builds $builds

    #Write Output
    write-Info "Writing output file [$output]."
    Set-Content $output $outputmarkdown
    get-content $output | out-string
        
    #write output to json
    write-Info "Writing output file [$output]."
    $json = $builds | convertto-Json -Depth 4
    set-content (join-path $outputpath $outputjson) $json
}

END {}

