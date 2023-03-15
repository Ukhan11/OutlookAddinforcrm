Office.initialize = function () {
    if (Office.context.mailbox.diagnostics.hostName === 'Outlook') {
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.AppointmentSend,
            onAppointmentSendHandler
        );
    }
}

function onAppointmentSendHandler(eventArgs) {
    eventArgs.preventDefault();

    // Get a callback token with REST permissions.
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            const accessToken = result.value;

            // Use the access token to get the current appointment and send it to CRM.
            getCurrentItem(accessToken);
        } else {
            console.log('Error getting callback token with REST permissions');
        }
    });
}

function getCurrentItem(accessToken) {
    // Get the item's REST ID.
    const itemId = getItemRestId();
    const getAppointmentUrl = `${Office.context.mailbox.restUrl}/v2.0/me/events/${itemId}`;

    $.ajax({
        url: getAppointmentUrl,
        dataType: 'json',
        headers: { 'Authorization': `Bearer ${accessToken}` }
    }).done(function (item) {
        // Appointment is passed in `item`.
        const appointment = item;

        // Send the appointment to CRM.
        sendAppointmentToCRM(appointment, accessToken);
    }).fail(function (error) {
        console.log(`Error getting appointment from Office REST API: ${error.responseText}`);
    });
}

function sendAppointmentToCRM(appointment, accessToken) {
    // Connect to CRM.
    const service = ConnectToMSCRM(
        'clientid',
        'client secret',
        'crm url',
        'tenantid',
        "accounts"
    );

    // Create a new CRM appointment from the Office appointment.
    const newAppointment = {
        subject: appointment.Subject,
        description: appointment.Body.Content,
        // Add other appointment properties as needed.
    };

    // Get the organizer of the appointment.
    const organizer = appointment.Organizer.EmailAddress.Name;

    // Store the subject and organizer in the account table in CRM.
    const table = {
        name: appointment.Subject,
        description: `Organized by ${organizer}`,
        // Add other fields as needed.
    };

    // Send the new appointment and account to CRM.
    service.Create(newAppointment, function (result) {
        console.log(`Appointment sent to CRM with ID ${result.id}`);

        // Get the ID of the new appointment.
        const appointmentId = result.id;

        // Set the account ID to the ID of the new appointment.
        table.accountid = appointmentId;

        service.Create(table, function (result) {
            console.log(`Account stored in CRM table with ID ${result.id}`);
        }, function (error) {
            console.log(`Error storing account in CRM table: ${error.message}`);
        });

    }, function (error) {
        console.log(`Error sending appointment to CRM: ${error.message}`);
    });
}




if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
}
