#[System.Reflection.Assembly]::Load("Sage.Peachtree.API.Resolver, Version=2015.0.0.932, Culture=neutral, PublicKeyToken=d06c16dde04d83e4, processorArchitecture=x86")
#[System.Reflection.Assembly]::Load("Sage.Peachtree.API, Version=2015.0.0.932, Culture=neutral, PublicKeyToken=d06c16dde04d83e4, processorArchitecture=x86")

[System.Reflection.Assembly]::LoadWithPartialName("Sage.Peachtree.API.Resolver")
[System.Reflection.Assembly]::LoadWithPartialName("Sage.Peachtree.API")

Write-Host "Sage 50 ShipGear write-back performance test" -ForegroundColor Green

Write-Host "Enter server name:" -ForegroundColor Green
[string]$serverName = [Console]::ReadLine()
#[string]$serverName = "DEV08"

[Sage.Peachtree.API.Resolver.AssemblyInitializer]::Initialize

$session = New-Object Sage.Peachtree.API.PeachtreeSession
$session.Begin("YOUR DEV TOKEN HERE")

[Sage.Peachtree.API.CompanyIdentifierList]$companyList = $session.CompanyList($serverName)
foreach ($i in $companyList) { Write-Host $i.CompanyName: $i.Path }

Write-Host "Enter company name:" -ForegroundColor Green
[string]$companyName = [Console]::ReadLine()
#[string]$companyName = "Stone Arbor Landscaping"

Write-Host "Enter invoice number:" -ForegroundColor Green
[string]$documentKey = [Console]::ReadLine()
#[string]$documentKey = "1000"

$guid = [Guid]::Empty
foreach ($i in $companyList) { if ($i.CompanyName -eq $companyName) { $company = $i; break; } }
Write-Host $company.Guid selected -ForegroundColor Cyan

#$paramList = New-Object 'System.Collections.Generic.Dictionary[string, object]'
#$authResult = $session.RequestAccess($company, $paramList)

$authResult = $session.RequestAccess($company)

if ($authResult = [Sage.Peachtree.API.AuthorizationResult]::Granted)
{
    #[Sage.Peachtree.API.Company]$comp = $session.Open($company, $paramList)
    [Sage.Peachtree.API.Company]$comp = $session.Open($company)
    
    [Sage.Peachtree.API.SalesInvoiceList]$list = $comp.Factories.SalesInvoiceFactory.List()
    
    [Sage.Peachtree.API.Collections.Generic.FilterExpression] $filter = [Sage.Peachtree.API.Collections.Generic.FilterExpression]::Equal([Sage.Peachtree.API.Collections.Generic.FilterExpression]::Property("SalesInvoice.ReferenceNumber"), [Sage.Peachtree.API.Collections.Generic.FilterExpression]::Constant($documentKey))
    $modifiers = [Sage.Peachtree.API.Collections.Generic.LoadModifiers]::Create()
    $modifiers.Filters = $filter

    $list.Load($modifiers)
    Write-Host Invoices selected: $list.Count

    $enum = $list.GetEnumerator()
    $enum.MoveNext()

    $cust = $comp.Factories.CustomerFactory.Load($enum.Current.CustomerReference)
    $enum.Current.CustomerNote = "Test Customer Note"
    $enum.Current.InternalNote = "Test Internal Note"
    $enum.Current.FreightAmount = "3.14"

    $w = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "Starting performance counter" -ForegroundColor Yellow
    $enum.Current.Save()
    $w.Stop()
    $elapsed=$w.Elapsed.TotalMilliseconds.ToString()
    Write-Host "Elapsed msec: $elapsed" -ForegroundColor Yellow

    Write-Host "Operation completed" -ForegroundColor Green
}
elseif ($authResult = [Sage.Peachtree.API.AuthorizationResult]::Pending)
{
    Write-Host "Authorization result: Pending - cannot continue" -ForegroundColor Red
}
else
{
    Write-Host "Authorization result: $authResult - cannot continue" -ForegroundColor Red
}

$session.End()