.DEFAULT_GOAL := run
.PHONY: build clean restore watch run

run:
	dotnet run --project src/bhmt.examples.excel.api/bhmt.examples.excel.api.csproj
build:
	dotnet build
clean:
	dotnet clean
restore:
	dotnet restore
watch:
	dotnet watch --project src/bhmt.examples.excel.api/bhmt.examples.excel.api.csproj run