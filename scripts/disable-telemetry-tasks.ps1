# Disables telemetry, feedback, and promo scheduled tasks. Run from an elevated PowerShell.
# Reverse any line with: Enable-ScheduledTask -TaskPath <path> -TaskName <name>
$t = @(
'\Microsoft\Windows\Customer Experience Improvement Program\Consolidator',
'\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip',
'\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser Exp',
'\Microsoft\Windows\Feedback\Siuf\DmClient',
'\Microsoft\Windows\Feedback\Siuf\DmClientOnScenarioDownload',
'\Microsoft\Windows\Device Information\Device',
'\Microsoft\Windows\Device Information\Device User',
'\Microsoft\Windows\Windows Error Reporting\QueueReporting',
'\Microsoft\Windows\Flighting\FeatureConfig\UsageDataFlushing',
'\Microsoft\Windows\Flighting\FeatureConfig\UsageDataReceiver',
'\Microsoft\Windows\Flighting\FeatureConfig\UsageDataReporting',
'\Microsoft\Windows\Sustainability\SustainabilityTelemetry',
'\Microsoft\Windows\Shell\ThemesSyncedImageDownload',
'\Microsoft\Windows\Shell\FamilySafetyMonitor',
'\Microsoft\Windows\Shell\FamilySafetyRefreshTask',
'\Microsoft\Windows\AccountHealth\RecoverabilityToastTask',
'\Microsoft\Windows\PerformanceTrace\ShowFeedbackToast',
'\Microsoft\Windows\PerformanceTrace\WhesvcToast',
'\Microsoft\Windows\Diagnosis\RecommendedTroubleshootingScanner',
'\OneDrive Reporting Task-S-1-5-21-2940037537-759120422-2483470356-1107',
'\Microsoft\Office\Office Performance Monitor',
'\Microsoft\Office\Office Feature Updates',
'\Microsoft\Office\Office Feature Updates Logon'
)
foreach ($p in $t) {
  $d = Split-Path $p; $n = Split-Path $p -Leaf
  $tp = if ($d -eq '\') { '\' } else { $d + '\' }
  try { Disable-ScheduledTask -TaskPath $tp -TaskName $n -ErrorAction Stop | Out-Null; "disabled $p" }
  catch { "FAILED  $p : $($_.Exception.Message)" }
}
