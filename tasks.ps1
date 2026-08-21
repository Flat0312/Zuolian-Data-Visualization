param(
    [Parameter(Position = 0)]
    [ValidateSet("run", "build-static", "test", "lint")]
    [string]${Task} = "run",

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]${TaskArgs}
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Push-Location $RepoRoot
try {
    switch ($Task) {
        "run" {
            python -m streamlit run app.py @TaskArgs
        }
        "build-static" {
            python build_static_site.py @TaskArgs
        }
        "test" {
            python -m pytest @TaskArgs
        }
        "lint" {
            python -m ruff check app.py build_static_site.py kb_schema.py app research/analysis @TaskArgs
        }
    }
}
finally {
    Pop-Location
}
