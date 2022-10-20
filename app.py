
if __name__=='__main__':
    import win32com.client as client

    # Create Item
    outlook = client.Dispatch("outlook.application")
    reserva = outlook.CreateItem(1)

    # Inputs
    subject = input("Meeting's title: ")
    location = input("Meeting's location: ")
    required_input = input('Required: ')
    day = str(input("Meeting's day (for example: DD/MM/YYYY): "))
    hour = str(input("Meeting's hour (for example: 09:00:00 AM): "))
    time = float(input("Meeting's time (in hours): "))

    # Calculations
    duration = time * 60
    start = f"{day} {hour}"

    # Add informations about meeting
    reserva.subject = subject
    reserva.location = location
    required = reserva.Recipients.add(required_input)
    reserva.start = start
    reserva.duration = duration

    # Default value
    reserva.MeetingStatus = 1

    # Execute
    reserva.display()