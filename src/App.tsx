import React from 'react';
import { AuthenticationResult, AuthError, PublicClientApplication } from "@azure/msal-browser";
import ReactLoading from "react-loading";
import { service, factories, models, IEmbedConfiguration } from 'powerbi-client';
import "./App.css";
import * as config from "./Config";


let accessToken = "";
let embedUrl = "";
let datasetID = "";

let dateDebut;
let dateFin;
let debutInput: React.Ref<HTMLInputElement>;
let finInput: React.Ref<HTMLInputElement>;

let container: HTMLElement;
let refer: React.Ref<HTMLDivElement>;
const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);


interface AppProps { }
interface AppState { accessToken: string; embedUrl: string; error: string[]; datasetID: string; dateDebut: string; dateFin: string; }

class Bilban extends React.Component<AppProps, AppState>{

    constructor(value: AppProps) {
        super(value);
        this.state = { accessToken: "", embedUrl: "", error: [], datasetID: "", dateDebut: "", dateFin: "" };
        refer = React.createRef();
        debutInput = React.createRef();
        finInput = React.createRef();

        this.updateParameters = this.updateParameters.bind(this);

    }

    render(): JSX.Element {

        if (this.state.error.length) {

            this.state.error.forEach(line => {
                console.log(line)
                alert(line)
            })
        }
        else if (this.state.accessToken !== "" && this.state.embedUrl !== "") {
            const embedConfiguration: IEmbedConfiguration = {
                type: "report",
                tokenType: models.TokenType.Aad,
                accessToken,
                embedUrl,
                id: config.reportId,
                settings: {
                    panes: {
                        filters: {
                            visible: false
                        },
                        pageNavigation: {
                            visible: false
                        }
                    }
                }

            };

            const report = powerbi.embed(container, embedConfiguration);


            // Clear any other loaded handler events
            report.off("loaded");

            // Triggers when a content schema is successfully loaded
            report.on("loaded", function () {
                console.log("Report load successful");

                document.getElementById('parameters').classList.toggle("hidden");
            });

            // Clear any other rendered handler events
            report.off("rendered");

            // Triggers when a content is successfully embedded in UI
            report.on("rendered", function () {
                console.log("Report render successful");

            });

            // Clear any other error handler event
            report.off("error");

            // Below patch of code is for handling errors that occur during embedding
            report.on("error", function (event) {
                const errorMsg = event.detail;

                // Use errorMsg variable to log error in any destination of choice
                console.error(errorMsg);
            });
        }

        return (<>
            <div id="parameters" className="hidden">
                <div id="parameter">
                    <label htmlFor="dateDebut">Date Debut:</label>
                    <input type="date" id="name" name="dateDebut" ref={debutInput} defaultValue={this.state.dateDebut}></input>
                    <label htmlFor="dateFin">Date Fin:</label>
                    <input type="date" id="name" name="dateFin" ref={finInput} defaultValue={this.state.dateFin}></input>
                    <button onClick={this.updateParameters}></button>
                </div>
            </div>
            <div
                id="container"
                ref={refer}>
                <div id="loading">
                    <ReactLoading type="spin" color="#ffffff" height={50} width={50} />
                </div>
            </div>
        </>);


    }

    componentDidMount(): void {

        if (refer !== null) {
            container = refer["current"];
        }

        // User input - null check
        if (config.workspaceId === "" || config.reportId === "") {
            this.setState({ error: ["Please assign values to workspace Id and report Id in Config.ts file"] })
        } else {

            this.login();
        }
    }

    componentWillUnmount(): void {
        powerbi.reset(container);
    }

    login(): void {

        const thisObj = this;

        const msalConfig = {
            auth: {
                clientId: config.clientId
            }
        };

        const loginRequest = {
            scopes: config.scopes
        };

        const msalInstance = new PublicClientApplication(msalConfig);


        function handleResponse(response: AuthenticationResult): void {

            if (response !== null) {
                accessToken = response.accessToken;
                thisObj.setUsername(response.account.name);
                thisObj.tryRefreshUserPermissions();
                thisObj.getembedUrl();
            }
            else {
                const currentAccounts = msalInstance.getAllAccounts();

                if (currentAccounts.length === 0) {
                    msalInstance.loginRedirect(loginRequest);
                }
                else if (currentAccounts.length === 1) {
                    msalInstance.setActiveAccount(currentAccounts[0]);
                    thisObj.setUsername(currentAccounts[0].name);
                }
            }
        }

        msalInstance.handleRedirectPromise().then(handleResponse).catch((error: AuthError) => {
            this.setState({ error: ["Redirect error: " + error] });
        });

        if (msalInstance.getAllAccounts().length) {

            msalInstance.acquireTokenSilent(loginRequest).then(response => {
                accessToken = response.accessToken;
                this.setUsername(response.account.name);
                this.getembedUrl();
            }).catch((error: AuthError) => {
                if (error.name === "InteractionRequiredAuthError") {
                    msalInstance.acquireTokenRedirect(loginRequest);
                }
                else {
                    thisObj.setState({ error: [error.toString()] })
                }
            })

        }
        else {
            msalInstance.loginRedirect(loginRequest);

        }

    }

    tryRefreshUserPermissions(): void {
        fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "POST"
        })
            .then(function (response) {
                if (response.ok) {
                    console.log("User permissions refreshed successfully.");
                } else {
                    // Too many requests in one hour will cause the API to fail
                    if (response.status === 429) {
                        console.error("Permissions refresh will be available in up to an hour.");
                    } else {
                        console.error(response);
                    }
                }
            })
            .catch(function (error) {
                console.error("Failure in making API call." + error);
            });
    }

    getembedUrl(): void {
        const thisObj: this = this;

        fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/reports/" + config.reportId, {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "GET"
        })
            .then(function (response) {
                const errorMessage: string[] = [];
                errorMessage.push("Error occurred while fetching the embed URL of the report")
                errorMessage.push("Request Id: " + response.headers.get("requestId"));

                response.json()
                    .then(function (body) {
                        // Successful response
                        if (response.ok) {
                            embedUrl = body["embedUrl"];
                            datasetID = body["datasetId"];
                            thisObj.setState({ accessToken: accessToken, embedUrl: embedUrl });
                            thisObj.getParameters();
                        }
                        // If error message is available
                        else {
                            errorMessage.push("Error " + response.status + ": " + body.error.code);

                            thisObj.setState({ error: errorMessage });
                        }

                    })
                    .catch(function () {
                        errorMessage.push("Error " + response.status + ":  An error has occurred");

                        thisObj.setState({ error: errorMessage });
                    });
            })
            .catch(function (error) {

                // Error in making the API call
                thisObj.setState({ error: error });
            })
    }

    getParameters(): void {

        const thisObj: this = this;
        let trans;

        fetch("https://api.powerbi.com/v1.0/myorg/datasets/" + datasetID + "/parameters", {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "GET"
        }).then(function (response) {
            const errorMessage: string[] = [];
            errorMessage.push("Error occurred while fetching the Paremeters of the report")
            errorMessage.push("Request Id: " + response.headers.get("requestId"));
            response.json()
                .then(function (body) {
                    if (response.ok) {

                        console.log(accessToken);

                        dateDebut = body['value'][0]['currentValue'];
                        dateDebut = dateDebut.split('/')[2] + '-' + dateDebut.split('/')[1] + '-' + dateDebut.split('/')[0];
                        trans = new Date(dateDebut)
                        dateDebut = trans.toISOString().split('T')[0]

                        dateFin = body['value'][1]['currentValue'];
                        dateFin = dateFin.split('/')[2] + '-' + dateFin.split('/')[1] + '-' + dateFin.split('/')[0];
                        trans = new Date(dateFin)
                        dateFin = trans.toISOString().split('T')[0]

                        thisObj.setState({ dateDebut: dateDebut, dateFin: dateFin });
                    }
                    else {
                        errorMessage.push("Error " + response.status + ": " + body.error.code);

                        thisObj.setState({ error: errorMessage });
                    }
                })
                .catch(function () {
                    errorMessage.push("Error " + response.status + ":  An error has occurred");

                    thisObj.setState({ error: errorMessage });
                });
        })
            .catch(function (error) {

                // Error in making the API call
                thisObj.setState({ error: error });
            })
    }

    async updateParameters(event) {

        let init = new Date(debutInput['current'].value);
        var Invetaire = new Date(init.getTime());
        Invetaire.setDate(init.getDate() - 1);

        const thisObj: this = this;

        await fetch("https://api.powerbi.com/v1.0/myorg/datasets/" + datasetID + "/Default.UpdateParameters", {
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json",
                "Content-Type": "application/json"
            },
            body: JSON.stringify({

                "updateDetails": [
                    {
                        "name": "DateDebut",
                        "newValue": new Date(debutInput['current'].value).toLocaleDateString('fr-FR')
                    },
                    {
                        "name": "DateFin",
                        "newValue": new Date(finInput['current'].value).toLocaleDateString('fr-FR')
                    },
                    {
                        "name": "DebutDePeriod",
                        "newValue": Invetaire.toLocaleDateString('fr-FR')
                    }
                ]

            }),
            method: "POST"
        })
            .then(function (response) {
                const errorMessage: string[] = [];
                errorMessage.push("Error occurred while upadting the parameters")
                errorMessage.push("Request Id: " + response.headers.get("requestId"));

                if (response.ok) {
                    console.log(response);
                    thisObj.refreshDataset();
                }
                else {
                    errorMessage.push("Error " + response.status + ": " + response.statusText);

                    thisObj.setState({ error: errorMessage });
                }


            })
            .catch(function (error) {
                // Error in making the API call
                console.log(error);
                thisObj.setState({ error: error });
            });

    }

    async refreshDataset() {

        const thisObj: this = this;

        await fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/datasets/" + datasetID + "/refreshes", {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "POST"
        })
            .then(function (response) {
                const errorMessage: string[] = [];
                errorMessage.push("Error occurred while upadting the parameters")
                errorMessage.push("Request Id: " + response.headers.get("requestId"));

                if (response.ok) {
                    console.log(response);
                    alert('all ok');
                    thisObj.getDataRefreshStatus();
                }
                else {
                    errorMessage.push("Error " + response.status + ": " + response.statusText);

                    thisObj.setState({ error: errorMessage });
                }

            })
            .catch(function (error) {
                // Error in making the API call
                console.log(error);
                thisObj.setState({ error: error });
            });
    }

    async getDataRefreshStatus() {

        const thisObj: this = this;

        await fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/datasets/" + datasetID + "/refreshes?$top=1", {
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            method: "GET"
        }).then(async function (response) {
            const errorMessage: string[] = [];
            errorMessage.push("Error");
            errorMessage.push("Request Id: " + response.headers.get("requestId"));
            await response.json()
                .then(function (body) {
                    if (response.ok) {
                        if (body.value[0]['status'] === 'Unknown') {
                            console.log('3iw')
                            setInterval(function () {
                                thisObj.getDataRefreshStatus();
                            }
                                , 30000
                            )
                        }
                        if (body.value[0]['status'] === 'Completed') {
                            alert('Update done refresh page')
                            window.location.reload();
                        }
                    }
                    else {
                        errorMessage.push("Error " + response.status + ": " + body.error.code);

                        thisObj.setState({ error: errorMessage });
                    }
                })
                .catch(function () {
                    errorMessage.push("Error " + response.status + ":  An error has occurred");

                    thisObj.setState({ error: errorMessage });
                });
        })
            .catch(function (error) {

                // Error in making the API call
                thisObj.setState({ error: error });
            })
    }

    setUsername(username: string): void {
        const welcome = document.getElementById("welcome");
        if (welcome !== null)
            welcome.innerText = "Bienvenu, " + username;
    }
}

export default Bilban;


