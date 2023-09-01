import React, { useState } from 'react'

export default function Home() {

    let [loader, setLoader] = useState(false);
    let [roomName, setRoomName] = useState('');

    const addBody = () => {
        setLoader(true);
        generate_access_token();
    };

    // generate access token
    function generate_access_token() {

        var myHeaders = new Headers();
        myHeaders.append("Content-Type", "application/json");

        var raw = JSON.stringify({
            "email": "voxotest11@gmail.com",
            "password": "iRRQuvrKgsoT"
        });

        var requestOptions = {
            method: 'POST',
            headers: myHeaders,
            body: raw,
            redirect: 'follow'
        };

        fetch("https://api.voxo.co/v2/authentication", requestOptions)
            .then(response => response.json())
            .then((result) => {

                const createMeetData = {
                    access_Token: result.accessToken,
                    extNum: result.user.extNum,
                    tenantId: result.user.tenantId,
                    userId: result.user.id,
                }
                create_meeting(createMeetData)

            })
            .catch((error) => {
                console.log('error: ', error);
                setLoader(false);
            });
    };

    // generate meeting info using api
    function create_meeting(meetData) {

        const token = meetData.access_Token

        var myHeaders = new Headers();
        myHeaders.append("Content-Type", "application/json");
        myHeaders.append("Authorization", `Bearer ${token}`);


        var raw = JSON.stringify({
            "type": "onDemand",
            "ext": meetData.extNum,
            "tenantId": meetData.tenantId,
            "userId": meetData.userId,
            "roomName": roomName
        });

        var requestOptions = {
            method: 'POST',
            headers: myHeaders,
            body: raw,
            redirect: 'follow'
        };

        fetch("https://api.voxo.co/v2/meet/create", requestOptions)
            .then(response => response.json())
            .then((result) => {

                const meeting_info = {
                    roomId: result.roomId,
                    dialInPin: result.dialInPin
                }

                insert_data(meeting_info);
                setRoomName('')
                setLoader(false);
            })
            .catch((error) => {
                console.log('error', error)
                setLoader(false);
            });
    }

    // insert data in body
    function insert_data(info) {
        const meetingInfo = `
            <p>To access the video meeting, simply click on the provided link: <a href="https://cx.voxo.co/meet/${info.roomId}">https://cx.voxo.co/meet/${info.roomId}</a></p>
            <p>Alternatively, if you prefer to join via phone, dial 4152124409 and enter pin ${info.dialInPin}.</p>
            <p>For sharing purposes, the room ID is <i>${info.roomId}</i>.</p>
            `;

        const item = Office.context.mailbox.item;
        item.body.setAsync(meetingInfo, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Body set successfully!");
                setLoader(false)
            } else {
                console.error(`Failed to set body with error: ${result.error.message}`);
                setLoader(false)
            }
        });
    }

    const isInputValid = roomName.trim().length > 0;

    return (
        <div className='container'>
            <p>The VOXO meetings add-in allow you to add VOXO meetings to email and calendar invites.</p>

            <input value={roomName} onChange={(e) => { setRoomName(e.target.value) }} type='text' placeholder='Meeting name...' className='room-input' />

            <button disabled={!isInputValid} onClick={addBody} className={`add-meeting-btn ${!isInputValid ? 'disabled' : ''}`}>Add Meeting info</button>

            {/* <div className='center-text'>
                <p>VOXO International numbers</p>
            </div> */}

            {loader ?
                <div className="loader-overlay">
                    <div className="loader"></div>
                </div> : ''
            }

            <div className='footer'>
                <p>© 2023 All rights reserved</p>
            </div>
        </div>
    )
}
