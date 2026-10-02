# officecorpus adapter: Open XML SDK (DocumentFormat.OpenXml 3.5.1)
FROM mcr.microsoft.com/dotnet/sdk:9.0.318 AS build
ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /src
COPY OfficeCorpusOpenXml.csproj ./
RUN dotnet restore OfficeCorpusOpenXml.csproj --source https://api.nuget.org/v3/index.json
COPY Program.cs ./
RUN dotnet publish OfficeCorpusOpenXml.csproj --configuration Release --no-restore --output /out

FROM mcr.microsoft.com/dotnet/runtime:9.0.20
LABEL org.officecorpus.adapter="open-xml-sdk" \
      org.officecorpus.library="DocumentFormat.OpenXml" \
      org.officecorpus.library-version="3.5.1" \
      org.officecorpus.runtime=".NET 9.0.20"
ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /app
COPY --from=build /out/ ./
USER $APP_UID
ENTRYPOINT ["dotnet", "/app/OfficeCorpusOpenXml.dll"]
