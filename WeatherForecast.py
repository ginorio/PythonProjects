import requests

# Set the API key
api_key = "4081103fa088c592aca6e065e4ff5657"

# Get the user input
location = input("Inserisci la località: ")

# Set the units
units = "metric"

# Make the request
response = requests.get("https://api.openweathermap.org/data/2.5/weather?q=" + location + "&appid=" + api_key + "&units=" + units)

# Check the response status code
if response.status_code == 200:

    # Get the response data
    weather_data = response.json()

    # Print the weather information
    print("Il tempo atmosferico a " + location + " è questo :")
    print("* Temperatura: " + str(weather_data["main"]["temp"]) + "°C")
    print("* Umidità: " + str(weather_data["main"]["humidity"]) + "%")
    print("* Pressione: " + str(weather_data["main"]["pressure"]) + "hPa")
    print("* Velocità del vento " + str(weather_data["wind"]["speed"]) + "m/s")
    print("* Nuvolosità: " + str(weather_data["clouds"]["all"]) + "%")
    print("* Descrizione: " + weather_data["weather"][0]["description"])

else:

    print("Error: " + str(response.status_code))

