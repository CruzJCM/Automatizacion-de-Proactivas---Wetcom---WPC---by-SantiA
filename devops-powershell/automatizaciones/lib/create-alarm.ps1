
Function New-Alarm ($esxi, $vcentername, $method, $action) {
    #---------------Content---------------
    $_this = Get-View -Id 'ServiceInstance-ServiceInstance'
    $_this.Content

    #---------------Description---------------
    $_this = Get-View -Id 'AlarmManager-AlarmManager'
    $_this.Description

    #---------------Content---------------
    $_this = Get-View -Id 'ServiceInstance-ServiceInstance'
    $_this.Content

    #---------------Description---------------
    $_this = Get-View -Id 'EventManager-EventManager'
    $_this.Description

    #---------------Content---------------
    $_this = Get-View -Id 'ServiceInstance-ServiceInstance'
    $_this.Content

    #---------------PerfCounter---------------
    $_this = Get-View -Id 'PerformanceManager-PerfMgr'
    $_this.PerfCounter

    #---------------Content---------------
    $_this = Get-View -Id 'ServiceInstance-ServiceInstance'
    $_this.Content

    #---------------CreateAlarm---------------
    $entity = New-Object VMware.Vim.ManagedObjectReference
    $entity.Type = 'HostSystem'
    $entity.Value = $h.Id -replace 'HostSystem-'
    $spec = New-Object VMware.Vim.AlarmSpec
    $spec.Expression = New-Object VMware.Vim.OrAlarmExpression
    $spec.Expression.Expression = New-Object VMware.Vim.AlarmExpression[] (1)
    $spec.Expression.Expression[0] = New-Object VMware.Vim.StateAlarmExpression
    $spec.Expression.Expression[0].Red = 'connected'
    $spec.Expression.Expression[0].Type = 'HostSystem'
    $spec.Expression.Expression[0].Operator = 'isEqual'
    $spec.Expression.Expression[0].StatePath = 'runtime.connectionState'
    $spec.Name = "falso positivo $vcentername"
    $spec.Action = New-Object VMware.Vim.GroupAlarmAction
    $spec.Action.Action = New-Object VMware.Vim.AlarmAction[] (1)
    $spec.Action.Action[0] = New-Object VMware.Vim.AlarmTriggeringAction
    $spec.Action.Action[0].TransitionSpecs = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec[] (1)
    $spec.Action.Action[0].TransitionSpecs[0] = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec
    $spec.Action.Action[0].TransitionSpecs[0].Repeats = $false
    $spec.Action.Action[0].TransitionSpecs[0].StartState = 'yellow'
    $spec.Action.Action[0].TransitionSpecs[0].FinalState = 'red'
    $spec.Action.Action[0].Yellow2green = $false
    $spec.Action.Action[0].Yellow2red = $false
    $spec.Action.Action[0].Red2yellow = $false
    switch ($method) {
        "Email" { 
            $spec.Action.Action[0].Action = New-Object VMware.Vim.SendEmailAction
            $spec.Action.Action[0].Action.Subject = 'Alarm {alarmName} on Host : {targetName} is {newStatus}'
            $spec.Action.Action[0].Action.CcList = ''
            $spec.Action.Action[0].Action.ToList = $action
            $spec.Action.Action[0].Action.Body = ''
        }
        "Script" {
            $spec.Action.Action[0].Action = New-Object VMware.Vim.RunScriptAction
            $spec.Action.Action[0].Action.Script = $action
        }
        Default {}
    }
    
    $spec.Action.Action[0].Green2yellow = $false
    $spec.Description = ''
    $spec.Enabled = $true
    $spec.Setting = New-Object VMware.Vim.AlarmSetting
    $spec.Setting.ToleranceRange = 0
    $spec.Setting.ReportingFrequency = 300
    $_this = Get-View -Id 'AlarmManager-AlarmManager'
    $_this.CreateAlarm($entity, $spec)
}
