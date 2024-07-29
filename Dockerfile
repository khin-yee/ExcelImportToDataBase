# Base runtime image
FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

# Build stage
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /

# Copy project files and restore dependencies
COPY ["Excel_Database/Excel_Database.csproj", "Excel_Database/"]
COPY ["Excel_Database.Domain/Excel_Database.Domain.csproj", "Excel_Database.Domain/"]
COPY ["Excel_Databse.Repository/Excel_Databse.Repository.csproj", "Excel_Databse.Repository/"]
COPY ["Excel_Databse.Service/Excel_Databse.Service.csproj", "Excel_Databse.Service/"]


RUN dotnet restore "Excel_Database/Excel_Database.csproj"

# Copy the rest of the code and build the application
COPY . .
WORKDIR "Excel_Database"
RUN dotnet build "Excel_Database.csproj" -c Release -o /app/build

# Publish stage
FROM build AS publish
RUN dotnet publish "Excel_Database.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Final runtime image
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Excel_Database.dll"]
