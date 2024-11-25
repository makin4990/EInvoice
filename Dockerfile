# Use the official ASP.NET image as the base for runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER app
WORKDIR /app
EXPOSE 5005
EXPOSE 5005

# Use the official SDK image for building the application
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy the project files and restore dependencies
COPY ["EInvoice/EInvoice.csproj", "EInvoice/"]
RUN dotnet restore "EInvoice/EInvoice.csproj"

# Copy the rest of the source code and build the application
COPY . .
WORKDIR "/src/EInvoice"
RUN dotnet build "EInvoice.csproj" -c Release -o /app/build

# Publish the application in Release mode
FROM build AS publish
RUN dotnet publish "EInvoice.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Set the base image and copy the published output
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .

# Set the entry point to run the application on the desired port
ENTRYPOINT ["dotnet", "EInvoice.dll"]

# Configure the application to listen on port 5449 by default
ENV ASPNETCORE_URLS=http://+:5448
