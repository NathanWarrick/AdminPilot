# TODO: CHECK THIS

from datetime import date, datetime
import win32com.client

import version

__version__ = version.version


def attendance_update(name, date_str, time_str, returning, reason, collected):
    # If date was blank insert todays date
    if date_str == "":
        date_str = date.today().strftime("%d/%m/%Y")

    # If time was blank insert current time
    if time_str == "":
        time_str = datetime.now().strftime("%H:%M")

    # If returning was blank insert No
    if returning == "":
        returning = "No"

    # Email is created and processed
    ol = win32com.client.Dispatch("Outlook.Application")
    mailItem = ol.CreateItem(0)
    mailItem.BodyFormat = 1
    mailItem.To = "absences@horsham-college.vic.edu.au"
    mailItem.Subject = "Attendance Update"
    if collected == "":
        mailItem.htmlBody = (
            """
            <h1>
            Name:
            """
            + str(name)
            + """
            <br><br>
            
            Date:
            """
            + str(date_str)
            + """
            <br><br>
            
            Time:
            """
            + str(time_str)
            + """
            <br><br>
            
            Returning? If so, what time?
            """
            + str(returning)
            + """
            <br><br>
            
            Reason:
            """
            + str(reason)
            + """
            <br><br>
            <br><br>
            <br><br>
            
            </h1>
            <p class="adminpilot">
                Sent with AdminPilot v
                """
            + __version__
            + """
                <br><br>
                https://github.com/NathanWarrick/AdminPilot
            </p>
            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            .adminpilot {
                text-shadow: 0px 0px;
                text-align: left;
                font-family: sans-serif;
                font-size: 15px;
                color: black;
                font-style: italic;
            }
            </style>
            """
        )
    else:
        mailItem.htmlBody = (
            """
            <h1>
            Name:
            """
            + str(name)
            + """
            <br><br>
            
            Date:
            """
            + str(date_str)
            + """
            <br><br>
            
            Time:
            """
            + str(time_str)
            + """
            <br><br>
            
            Returning? If so, what time?
            """
            + str(returning)
            + """
            <br><br>
            
            Reason:
            """
            + str(reason)
            + """
            <br><br>
            
            Who Collected:
            """
            + str(collected)
            + """
            <br><br>
            <br><br>
            <br><br>
            </h1>
            <p class="adminpilot">
                Sent with AdminPilot v
                """
            + __version__
            + """
                <br><br>
                https://github.com/NathanWarrick/AdminPilot
            </p>

            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            .adminpilot {
                text-shadow: 0px 0px;
                text-align: left;
                font-famsily: sans-serif;
                font-size: 15px;
                color: black;
                font-style: italic;
                }
            </style>
            """
        )
    # mailItem.Display()  # email is displayed prior to sending
    mailItem.Send()  # email is sent


def student_ID(name):
    # Email is created and processed
    ol = win32com.client.Dispatch("Outlook.Application")
    mailItem = ol.CreateItem(0)
    mailItem.BodyFormat = 1
    mailItem.To = "8818-helpdesk@schools.vic.edu.au"  # enter IT email here
    mailItem.Subject = "Student ID Request"
    mailItem.htmlBody = (
        """
        <p>
        Hi IT,
        <br><br>
        Can i please get a Student IT card made up for the following student as they have paid their $5 fee.
        </p>       
        <p class="bolded">
        """
        + str(name)
        + """
        </p>      
        <p>
        <br><br>
        Thank you!
        <br><br>   
        <br><br> 
        <br><br>  
        </p>
    
        <p class="adminpilot">
        Sent with AdminPilot v
        """
        + __version__
        + """
        <br><br>
        https://github.com/NathanWarrick/AdminPilot
        </p>
        
        <style>
        p {
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 18px;
            color: black;
            }
        .bolded {
            font-weight: bold;
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 20px;
            color: black;
            }
        <style>
        .adminpilot {
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 15px;
            color: black;
            font-style: italic;
            }
        </style>
        """
    )

    # mailItem.Display()  # email is displayed prior to sending
    mailItem.Send()  # email is sent
