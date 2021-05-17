#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:5.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:5.0 AS build
WORKDIR /src
COPY ["WebProject.csproj", "."]
RUN dotnet restore "./WebProject.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "WebProject.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "WebProject.csproj" -c Release -o /app/publish

ADD https://www.microsoft.com/en-gb/download/confirmation.aspx?id=7151
RUN Jet40SP8_9xNT.exe /Q

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "WebProject.dll"]