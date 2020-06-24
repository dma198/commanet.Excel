#!/bin/bash
dotnet build -c release 
nuget pack ./package/Package.nuspec
