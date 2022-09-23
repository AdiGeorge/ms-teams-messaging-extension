import {
    TeamsActivityHandler,
    TurnContext,
    MessageFactory,
    CardFactory,
    MessagingExtensionAction,
    MessagingExtensionActionResponse,
    MessagingExtensionAttachment
  } from "botbuilder";
  import {find, sortBy} from "lodash";
  
  import * as Util from "util";
  import * as debug from "debug";
  
  const TextEncoder = Util.TextEncoder;
  const log = debug("msteams");
  
  export class MessBot extends TeamsActivityHandler {
    protected handleTeamsMessagingExtensionFetchTask(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
      const planets: any = require("./planets.json");
      const sortedPlanets: any = sortBy(planets, ["id"])
        .map((planet) => {
          return { value: planet.id, title: planet.name };
        });
      const adaptiveCardSource: any = require("./planetSelectorCard.json");
      const planetChoiceSet: any = find(adaptiveCardSource.body, { id: "planetSelector" });
      planetChoiceSet.choices = sortedPlanets;
      const adaptiveCard = CardFactory.adaptiveCard(adaptiveCardSource);
    
      const response: MessagingExtensionActionResponse = {
        task: {
          type: "continue",
          value: {
            card: adaptiveCard,
            title: "Planet Selector",
            height: 150,
            width: 500
          }
        }
      } as MessagingExtensionActionResponse;
    
      return Promise.resolve(response);
    }

    protected handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
      switch (action.commandId) {
        case "planetExpanderAction": {
          const planets: any = require("./planets.json");
          const selectedPlanet: any = planets.filter((planet) => planet.id === action.data.planetSelector)[0];
          const adaptiveCard = this.getPlanetDetailCard(selectedPlanet);
    
          return Promise.resolve({
            composeExtension: {
              type: "result",
              attachmentLayout: "list",
              attachments: [adaptiveCard]
            }
          } as MessagingExtensionActionResponse);
        }
        default:
          throw new Error("NotImplemented");
      }
    }

    private getPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {

      const adaptiveCardSource: any = require("./planetDisplayCard.json");
    
      adaptiveCardSource.actions[0].url = selectedPlanet.wikiLink;
      find(adaptiveCardSource.body, { id: "cardHeader" }).items[0].text = selectedPlanet.name;
      const cardBody: any = find(adaptiveCardSource.body, { id: "cardBody" });
      find(cardBody.items, { id: "planetSummary" }).text = selectedPlanet.summary;
      find(cardBody.items, { id: "imageAttribution" }).text = "*Image attribution: " + selectedPlanet.imageAlt + "*";
      const cardDetails: any = find(cardBody.items, { id: "planetDetails" });
      cardDetails.columns[0].items[0].url = selectedPlanet.imageLink;
      find(cardDetails.columns[1].items[0].facts, { id: "orderFromSun" }).value = selectedPlanet.id;
      find(cardDetails.columns[1].items[0].facts, { id: "planetNumSatellites" }).value = selectedPlanet.numSatellites;
      find(cardDetails.columns[1].items[0].facts, { id: "solarOrbitYears" }).value = selectedPlanet.solarOrbitYears;
      find(cardDetails.columns[1].items[0].facts, { id: "solarOrbitAvgDistanceKm" }).value = Number(selectedPlanet.solarOrbitAvgDistanceKm).toLocaleString();
    
      return CardFactory.adaptiveCard(adaptiveCardSource);
    }
  }