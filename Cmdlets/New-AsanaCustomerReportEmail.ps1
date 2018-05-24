function New-AsanaCustomerReportEmail
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Report
    )

    Begin
    {
        $ol = New-Object -comObject Outlook.Application 
        $ns = $ol.GetNameSpace("MAPI")
    }
    Process
    {
        
        $template = '
            <html>
                <head>
                    <style type="text/css">
                        div#wrapper {
                            margin: 20px 20px 20px 20px;
                        }

                        table { 
                            border-spacing: 0;
                            border-collapse: collapse;
                            border-color: rgb(200,200,200);
                        }

                        tr.task {
                            min-height: 26px;
                            border: solid rgb(200,200,200);
                            border-width: 1px 0px 1px 0px;
                    
                        }

                        tr.subtask {
                            min-height: 26px;
                            border: solid rgb(200,200,200);
                            border-width: 1px 0px 1px 0px;
                    
                        }

                        td.text {
                            width: 800px;
                        }

                        tr.task td {
                            padding: 8px 2px 8px 4px;
                        }

                        tr.subtask td {
                            padding: 8px 2px 8px 4px;
                        }

                        tr.completed td {
                            color: rgb(200,200,200);
                        }

                        tr.phasecompleted td {
                            color: rgb(200,200,200);
                        }

                        tr.phase td {font-weight: bold;}


                        .icon {
                            fill: rgb(200,200,200);
                            color: rgb(200,200,200);
                            width: 12px;
                            height: 12px;
                        }

                        .incomplete .icon {
                            visibility: hidden;
                        }

                        html {
                            font-family: arial;
                        }

                        h1 {
                            font-size: 24px;
                        }

                        a.tag {
                            border: 1px solid rgb(200,200,200);
                            padding: 3px;
                            font-size: 10px;
                            border-radius: 4px;
                        }

                        a.subtasktext {
                            margin-left: 14px;
                        }
                    </style>
                </head>
                <body><h1>HEADERTEXT</h1><div id="wrapper"><table>TASKLIST</table></div></body>
            </html>
        '

        

        
        Write-Verbose "Working on $report"
        $project = $null
        $project = Get-AsanaProjectWithTasksAndPhase -ProjectNumber $report.AsanaProjectNumber -Verbose:$false -IncludeSubTasks:(!!$report.IncludeSubTasks)

        $tasklist = $project.Tasks | ? {$_.Task.name -cnotlike "Internal*"} | where {!$report.IncompleteTasksOnly -or !$_.Task.Completed}| foreach {
            $classes = @("task")
            if($_.Task.name -like "*:") {
                $classes += "phase"
                if($_.Task.Completed) {
                    $classes += "phasecompleted"
                } else {
                    $classes += "phaseincomplete"
                }
                "<tr class='$classes'><td colspan='3' >$($_.Task.name)</td></tr>"
            } else {
                $icon = '<svg class="icon" viewBox="0 0 32 32"><polygon points="27.672,4.786 10.901,21.557 4.328,14.984 1.5,17.812 10.901,27.214 30.5,7.615 "></polygon></svg>'
                [int] $percent = $_.task.custom_fields | ? id -eq 230589792216219 | select -exp number_value
                if($_.Task.due_on -and !$_.Task.Completed) {
                    $due = $_.Task.due_on
                } else {
                    $due = "&nbsp;"
                }

                if($_.Task.Completed) {
                    $classes += "completed"
                    $percent = 100
                } else {
                    $classes += "incomplete"
                }

                $tags = $_.task.tags| foreach {if($_.name) {"<a class='tag'>$($_.Name)</a>"}}

                "<tr class='$classes'><td class='text'>$icon <a class='tasktext'>$($_.Task.name) ($percent %)</a></td><td class='due'>$due</td><td class='tags'>$tags</td></tr>"

                if($_.Subtasks -and !$_.Task.Completed) {
                    $_.Subtasks | Foreach {
                    $classes = @("subtask")
                        $icon = '<svg class="icon" viewBox="0 0 32 32"><polygon points="27.672,4.786 10.901,21.557 4.328,14.984 1.5,17.812 10.901,27.214 30.5,7.615 "></polygon></svg>'
                        [int] $percent = $_.task.custom_fields | ? id -eq 230589792216219 | select -exp number_value
                        if($_.Task.due_on -and !$_.Task.Completed) {
                            $due = $_.Task.due_on
                        } else {
                            $due = "&nbsp;"
                        }

                        if($_.Task.Completed) {
                            $classes += "completed"
                            $percent = 100
                        } else {
                            $classes += "incomplete"
                        }

                        $tags = $_.task.tags| foreach {if($_.name) {"<a class='tag'>$($_.Name)</a>"}}

                        "<tr class='$classes'><td class='text'>$icon <a class='subtasktext'>$($_.Task.name) ($percent %)</a></td><td class='due'>$due</td><td class='tags'>$tags</td></tr>"
                    }
                }
            }
        }

        $result = $template -creplace "TASKLIST", ($tasklist -join "`n") -creplace "HEADERTEXT", $report.HeaderText
        $folder = "$($ENV:TEMP)\projectreport-$([guid]::NewGuid())" 
        mkdir $folder | Out-Null
        $file = "$folder\Report.html" 
        Set-Content -Path $file -Value $result -Encoding UTF8 
    
        $mail = $ol.CreateItem(0)
        $mail.Display() | Out-Null
        $mail.Subject = $report.Subject 
        $mail.Attachments.Add($file) | Out-Null
        $report.Recipients | foreach {
            $Mail.Recipients.Add($_) | Out-Null
        }
    
        [regex]$pattern = "<o:p>&nbsp;</o:p>"
        $mail.HTMLBody = $pattern.replace($mail.HTMLBody , "<o:p>$($report.MailText)</o:p>", 1) 
        


    }
    End
    {
        $ol = $null 
    }
}