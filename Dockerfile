#build container
FROM mcr.microsoft.com/dotnet/sdk:5.0-buster-slim as build

WORKDIR /src
COPY . .
RUN dotnet restore
RUN dotnet test --logger:trx
RUN dotnet publish -c Release -o /app

#runtime container
FROM mcr.microsoft.com/dotnet/aspnet:5.0.0-alpine3.12
RUN apk add --no-cache tzdata
RUN apk add icu-libs
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false

WORKDIR /app
COPY --from=build /app .
EXPOSE 5000
ENTRYPOINT ["dotnet", "testexcel.dll"]