import React from 'react';
import ReactDOM from 'react-dom/client';
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
let Site;
let debutInput: React.Ref<HTMLInputElement>;
let finInput: React.Ref<HTMLInputElement>;
let siteInput: React.Ref<HTMLSelectElement>;
let container: HTMLElement;
let refer: React.Ref<HTMLDivElement>;
const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

interface AppProps { }
interface AppState { accessToken: string; embedUrl: string; error: string[]; datasetID: string; dateDebut: string; dateFin: string; site: string }

class Bilban extends React.Component<AppProps, AppState>{

    constructor(value: AppProps) {
        super(value);
        this.state = { accessToken: "", embedUrl: "", error: [], datasetID: "", dateDebut: "", dateFin: "", site: "" };
        refer = React.createRef();
        debutInput = React.createRef();
        finInput = React.createRef();
        siteInput = React.createRef();
        this.handleChange = this.handleChange.bind(this);
        this.updateParameters = this.updateParameters.bind(this);

    }

    render(): JSX.Element {

        if (this.state.error.length) {

            this.state.error.forEach(line => {
                console.log(line)
                if (line.split(":")[0] === "BrowserAuthError") {
                    this.login();
                }
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
                document.getElementById('parameters')!.classList.toggle("hidden");
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
                    <label htmlFor="Site">Agence:</label>
                    <select name="Site" id="select" ref={siteInput} value={this.state.site} onChange={this.handleChange}>
                        <option value="0101 - Casablanca" id='0101 '>0101 - Casablanca</option>
                        <option value="0102 - FBS" id='0102 '>0102 - FBS</option>
                        <option value="0103 - EL Jadida" id='0103 '>0103 - EL Jadida</option>
                        <option value="0201 - Oujda" id='0201 '>0201 - Oujda</option>
                        <option value="0202 - Fes" id='0202 '>0202 - Fes  </option>
                        <option value="0203 - Meknes" id='0203 '>0203 - Meknes</option>
                        <option value="0301 - Marrakech" id='0301 '>0301 - Marrakech</option>
                        <option value="0303 - Kelaa" id='0303 '>0303 - Kelaa</option>
                        <option value="0304 - Safi" id='0304 '>0304 - Safi</option>
                        <option value="0305 - Essaouira" id='0305 '>0305 - Essaouira</option>
                        <option value="0501 - Rabat" id='0501 '>0501 - Rabat</option>
                        <option value="0502 - Tanger" id='0502 '>0502 - Tanger</option>
                        <option value="0503 - Souk Larbaa" id='0503 '>0503 - Souk Larbaa</option>
                    </select>
                    <button onClick={this.updateParameters}>Actualisez </button>
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

    handleChange(event) { this.setState({ site: event.target.value }); document.getElementById('parameters')!.classList.toggle("hidden"); }

    componentDidMount(): void {

        if (refer !== null) {
            container = refer["current"];
        }

        // User input - null check
        if (config.workspaceId === "" || config.reportId === "") {
            this.setState({ error: ["Please assign values to workspace Id and report Id in Config.ts file"] })
        }
        else {

            this.login();
        }
    }

    componentWillUnmount(): void {
        powerbi.reset(container);
    }

    async login(): Promise<void> {

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
                thisObj.setUsername(response.account!.name as unknown as string);
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
                    thisObj.setUsername(currentAccounts[0].name as unknown as string);
                }
            }
        }

        await msalInstance.handleRedirectPromise().then(() => handleResponse).catch((error: AuthError) => {
            this.setState({ error: ["Redirect error: " + error] });
        });

        if (msalInstance.getAllAccounts().length) {
            const currentAccounts = msalInstance.getAllAccounts();
            msalInstance.setActiveAccount(currentAccounts[0]);
            msalInstance.acquireTokenSilent(loginRequest).then(response => {
                accessToken = response.accessToken;
                this.setUsername(response.account!.name as unknown as string);
                this.getembedUrl();
            }).catch((error: AuthError) => {
                console.log(error.name);
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
                        console.log(body);
                        dateDebut = body['value'][0]['currentValue'];
                        dateDebut = dateDebut.split('/')[2] + '-' + dateDebut.split('/')[1] + '-' + dateDebut.split('/')[0];
                        trans = new Date(dateDebut)
                        dateDebut = trans.toISOString().split('T')[0]

                        dateFin = body['value'][1]['currentValue'];
                        dateFin = dateFin.split('/')[2] + '-' + dateFin.split('/')[1] + '-' + dateFin.split('/')[0];
                        trans = new Date(dateFin)
                        dateFin = trans.toISOString().split('T')[0]

                        Site = body['value'][3]['currentValue'];


                        thisObj.setState({ dateDebut: dateDebut, dateFin: dateFin, site: Site });
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

        let init = new Date(debutInput!['current'].value);
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
                        "newValue": new Date(debutInput!['current'].value).toLocaleDateString('fr-FR')
                    },
                    {
                        "name": "DateFin",
                        "newValue": new Date(finInput!['current'].value).toLocaleDateString('fr-FR')
                    },
                    {
                        "name": "DateInventaire",
                        "newValue": Invetaire.toLocaleDateString('fr-FR')
                    },
                    {
                        "name": "Site",
                        "newValue": siteInput!['current'].value as unknown as string
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
                    alert("l'actualisation a été effectuée avec succès. Les données sont maintenant mises à jour cela prendra plus de 10 minutes. Veuillez être patient. :)");
                    let root = ReactDOM.createRoot(document.getElementById('main') as HTMLElement)

                    root.render(<div id="load">
                        <p id="loadText">actualisation en cours....</p>
                        <div id='spin'>
                            <ReactLoading type="spin" color="#ffffff" height={50} width={50} />
                        </div>
                    </div>)
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
                            setInterval(function () {
                                thisObj.getDataRefreshStatus();
                            }
                                , 120000
                            )
                        }
                        if (body.value[0]['status'] === 'Completed') {
                            alert('Actualisation effectuée')
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


