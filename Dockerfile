# Use the official ASP.NET Core runtime image as a base
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

# Use the official .NET SDK image to build the app
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy the .csproj file and restore dependencies
COPY ["WebDataSortingJ.csproj", "./"]
RUN dotnet restore "WebDataSortingJ.csproj"

# Copy the rest of the application and build
COPY . .
WORKDIR "/src"
RUN dotnet build "WebDataSortingJ.csproj" -c Release -o /app/build

# Publish the application
FROM build AS publish
RUN dotnet publish "WebDataSortingJ.csproj" -c Release -o /app/publish

# Create final runtime image
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "WebDataSortingJ.dll"]
