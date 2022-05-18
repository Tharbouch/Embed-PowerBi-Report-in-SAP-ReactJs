import React from 'react';
import ReactDOM from 'react-dom';
import { AuthenticationResult, AuthError, PublicClientApplication } from "@azure/msal-browser";
import ReactLoading from "react-loading";
import { service, factories, models, IEmbedConfiguration } from 'powerbi-client';
import "./App.css";
import * as config from "./Config";


let accessToken = "";
let embedUrl = "";
let datasetID = "";
let container: HTMLElement;
let refer: React.Ref<HTMLDivElement>;
let loading: JSX.Element;
const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);


interface AppProps { }
interface AppState { accessToken: string; embedUrl: string; error: string[]; datasetID: string }

class Bilban extends React.Component<AppProps, AppState>{

    constructor(value: AppProps) {
        super(value);
        this.state = { accessToken: "", embedUrl: "", error: [], datasetID: "" };
        refer = React.createRef();

        loading = (
            <div
                id="container"
                ref={refer} >
                <div id="loading">
                    <ReactLoading type="spin" color="#ffffff" height={50} width={50} />
                </div>
            </div>

        )
    }

    render(): JSX.Element {

        const thisObj = this;

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
            });

            // Clear any other rendered handler events
            report.off("rendered");

            // Triggers when a content is successfully embedded in UI
            report.on("rendered", function () {
                console.log("Report render successful");



                document.getElementById('parameters').classList.toggle("hidden");

                const elem = React.createElement('div', {
                    id: 'parameter'
                }, [
                    <>
                        <label htmlFor="dateDebut" >Date Debut:</label>
                        <input type="date" id="name" name="dateDebut"  ></input>
                        <label htmlFor="dateFin" >Date Fin:</label>
                        <input type="date" id="name" name="dateFin"  ></input>
                        <label htmlFor="DateInventaire" >Date Inventaire:</label>
                        <input type="date" id="name" name="DateInventaire"  ></input>
                    </>
                ]
                );
                ReactDOM.render(elem, document.getElementById('parameters'))

                thisObj.getParameters();
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

        return loading;

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


