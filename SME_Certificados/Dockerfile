#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/runtime:6.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["SME_Certificados/SME_Certificados.csproj", "SME_Certificados/"]
RUN dotnet restore "SME_Certificados/SME_Certificados.csproj"
COPY . .
WORKDIR "/src/SME_Certificados"
RUN dotnet build "SME_Certificados.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "SME_Certificados.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet","SME_Certificados.dll"]
