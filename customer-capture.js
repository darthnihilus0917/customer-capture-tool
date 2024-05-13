const readline = require("readline");
const dotenv = require("dotenv");
dotenv.config();

const { processes, meats } = require('./lib/options/options');
const { loadTitle } = require('./lib/utils/utils');
const { appLabels } = require('./lib/contants/contants');

const { consolidateRobinson, consolidateMetro, 
    consolidatePuregold } = require('./lib/processes/consolidate');

const { buildPorkmeat, buildPoultry, 
    buildSwine } = require('./lib/processes/buildSOTC');

const { generatePorkmeat, generatePoultry, 
    generateSwine } = require("./lib/processes/generateDataSource");

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

function askQuestion(question, options) {
    return new Promise((resolve, reject) => {
        const numberedOptions = options.map((option, index) => `[${index + 1}] ${option}`);
        rl.question(question + "\n" + numberedOptions.join("\n") + "\n", (answer) => {
                const selectedOption = options[parseInt(answer) - 1];
                if (selectedOption) {
                    resolve(selectedOption.toUpperCase());
                } else {
                    console.log(appLabels.invalidAnswer);
                    askQuestion(question, options).then(resolve).catch(reject);
                }
            }
        );
    });
}

async function main() {
    try {
        let meat = "";
  
        while (meat !== "EXIT") {
            const meatOptions = meats;
    
            loadTitle();
            
            meat = await askQuestion("Select A Meat Data Source:", meatOptions);
    
            if (meat === "EXIT") {
                const confirmation = await askQuestion(appLabels.confirmExit,["Yes", "No"]);
                if (confirmation === "NO") {
                    meat = ""; // Reset store to continue the loop
                    continue;
                }
                console.log(appLabels.closingApp);
                rl.close();
                return;
            }
    
            console.log('\nYou selected:', meat);
    
            let actions = processes;
            let action = "";

            while (action !== "EXIT") {
                action = await askQuestion("\nWhat do you want to do?", actions);
  
                if (action === "EXIT") {
                    const confirmation = await askQuestion(appLabels.confirmExit, ["Yes", "No"]);
                    if (confirmation === "NO") {
                        action = ""; // Reset action to continue the loop
                        continue;
                    }
                    console.log(appLabels.closingApp);
                    rl.close();
                    return;
                }
                console.log('\nYou selected:', action);
  
                if (action === "CANCEL") {
                    break; // break to go back to meat selection
                }

                if (action === "COPY SOTC DATA") {
                    console.log(`Copying ${meat} SOTC & PICKUP data. Please wait...`);
                    switch(meat) {
                        case "PORK MEATS":
                            await buildPorkmeat(meat, action);
                            break;                    
                        case "POULTRY":
                            await buildPoultry(meat, action);
                            break;
                        case "SWINE":
                            await buildSwine(meat, action);
                            break;
                        default:
                            console.log(`${appLabels.processNotAvailable} ${meat}.`);
                    }
                }
      
                if (action === "GENERATE DATA SOURCE") {
                    console.log(`Generating ${meat} data source. Please wait...`);
                    switch(meat) {
                        case "PORK MEATS":
                            await generatePorkmeat(meat, action);
                            break;
                        case "POULTRY":
                            await generatePoultry(meat, action);
                            break;
                        case "SWINE":
                            await generateSwine(meat, action);
                            break;
                        default:
                            console.log(`${appLabels.processNotAvailable} ${meat}.`);
                    }
                }

                if (action === "CONSOLIDATE") {
                    console.log(`Consolidating ${meat} data. Please wait...`);
                    switch(meat) {
                        case "PORK MEATS":
                            await consolidateRobinson(meat, action);
                            break;                    
                        case "POULTRY":
                            await consolidatePuregold(meat, action);
                            break;
                        case "SWINE":
                            await consolidateMetro(meat, action);
                            break;
                        default:
                            console.log(`${appLabels.processNotAvailable} ${meat}.`);
                    }
                }                
            }
        }
    } catch (err) {
        console.error(err.message);
  
    } finally {
        rl.close();
    }
}
  
main();
  