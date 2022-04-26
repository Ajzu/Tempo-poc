#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

#Depending on the operating system of the host machines(s) that will build or run the containers, the image specified in the FROM statement may need to be changed.
#For more information, please see https://aka.ms/containercompat

FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["well-estimation-tool-grid-poc/well-estimation-tool-grid-poc.csproj", "well-estimation-tool-grid-poc/"]
RUN dotnet restore "well-estimation-tool-grid-poc/well-estimation-tool-grid-poc.csproj"
COPY . .
WORKDIR "/src/well-estimation-tool-grid-poc"
RUN dotnet build "well-estimation-tool-grid-poc.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "well-estimation-tool-grid-poc.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "well-estimation-tool-grid-poc.dll"]