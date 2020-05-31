(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    class room {
        constructor(tiles, doors, title, desc, it) {
            this.tiles = tiles;
            this.doors = doors;
            this.title = title;
            this.description = desc;
            this.items = it;
        }
    }

    //Tracking the player icon's position
    var playerPos = [3, 4];

    //Array list of room objects
    var roomList = [];

    //Tracks the current active room
    var currentRoom = null;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("Emergus Questivo. An adventure through the Portal Dimension of the Wizard weNnoR. Find three keys (ᚩ)");
                $('#button-text').text("weNnoR!");
                $('#button-desc').text("weNnoR!!!!");
                $('#highlight-button').click(roomRender);
                return;
            }

            //setCellSizes();
            roomList = makeRooms();
            //loadSampleData();
            currentRoom = roomList[0];
            moveRoom(currentRoom);

            $("#template-description").text("Navigate your way through the Wizard weNnoR's realm. Find three keys (ᚩ)");
            $('#button-text').text("Show me Potato Salad!");
            $('#button-desc').text("Highlights the largest number.");

            $('#btn-up-text').text("Up");
            $('#btn-up').click(function () {
                move('u', currentRoom);
            });

            $('#btn-down-text').text("Down");
            $('#btn-down').click(function () {
                move('d', currentRoom);
            });

            $('#btn-left-text').text("Left");
            $('#btn-left').click(function () {
                move('l', currentRoom);
            });

            $('#btn-right-text').text("Right");
            $('#btn-right').click(function () {
                move('r', currentRoom);
            });

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(roomRender);
            
        });
    };

    //Changing the current room.
    function moveRoom(newRoom) {
        Excel.run(function (ctx) {
            var currentSheet = ctx.workbook.worksheets.getActiveWorksheet();

            var sheets = ctx.workbook.worksheets;
            var newSheet = sheets.add(newRoom.title);
            sheets.load("items/name");
            newSheet.activate();
            currentSheet.delete();
            setCellSizes();
            roomRender(newRoom);
            currentRoom = newRoom;

            return ctx.sync();
        });
    }

    // Resizes the cells of the play area so that they are (more or less) square.
    function setCellSizes() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRanges("a1:k11");

            var internalRange = sheet.getRanges("c3:i9");
            cellRange.format.fill.color = "White";
            internalRange.format.fill.color = "yellow";
            cellRange.format.columnWidth = 20;
            cellRange.format.rowHeight = 20;
            cellRange.format.horizontalAlignment = "Center";
            cellRange.format.verticalAlignment = "Center";
            cellRange.format.font.size = 15;
            cellRange.format.font.color = "Black";
            return ctx.sync();
        }).catch(errorHandler);
    }

    function populateItems(newRoom)
    {
        console.log(newRoom);
        Excel.run(function (ctx)
        {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRange("c3:i9");
            cellRange.load("value");

            return ctx.sync(newRoom).then(function (newRoom)
            {
                console.log(newRoom.items);
                newRoom.items.forEach(function (itemEntry)
                {
                    console.log(itemEntry.row + ',' + itemEntry.col);
                    console.log(cellRange.getCell(itemEntry.row, itemEntry.col));
                    cellRange.getCell(itemEntry.row, itemEntry.col).values = itemEntry.item;
                });
            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    //Render rooms with a 2 cell pad on top and left sides (top left room edge starts at Row 3, Column C)
    function roomRender(newRoom) {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRange("b2:j10");
            cellRange.load("value, rowCount, columnCount");

            return ctx.sync().then(function () {
                for (var i = 0; i < cellRange.rowCount; i++) {
                    for (var j = 0; j < cellRange.columnCount; j++)
                    {
                        //Get reference to current cell
                        var currentCell = cellRange.getCell(i, j);

                        //Empties cell value during render
                        if (currentCell.value != '')
                        {
                            currentCell.values = '';
                        }

                        switch (newRoom.tiles[i][j]) {
                            case 0:
                                currentCell.format.fill.color = "Black";
                                break;
                            case 1:
                                currentCell.format.fill.color = "LightGrey"; //#D3D3D3
                                break;
                            case 2:
                                currentCell.format.fill.color = "AntiqueWhite"; //#FAEBD7
                                break;
                            case 3:
                                currentCell.format.fill.color = "DodgerBlue"; //#1E90FF
                                break;
                            case 4:
                                currentCell.format.fill.color = "SaddleBrown"; //#8B4513
                                break;
                            case 5:
                                currentCell.format.fill.color = "Yellow"; //#FFFF00
                                break;
                            case 6:
                                currentCell.format.fill.color = "Purple" //#800080
                                break;
                            default:
                                currentCell.format.fill.color = "black";
                                break;
                        }
                    }
                }

                //Display the player
                cellRange.getCell(playerPos[0], playerPos[1]).values = '☺';
            }).then(populateItems(newRoom)).then(ctx.sync);
                
        }).catch(errorHandler);
    }

    //Handles player movement checks and room transition checks
    //If tiles are "Floor" or "Sand" coloured, player icon move is applied
    //If tiles are a "Door" tile, move room is applied
    function move(direction, currentRoom) {
        var newPos = [-1, -1];
        var doorVal = -1;

        switch (direction) {
            case 'u':
                if (playerPos[0] > 0) {
                    newPos = [playerPos[0] - 1, playerPos[1]];
                    //playerPos = [7, 4];
                    doorVal = 0;
                }
                break;
            case 'd':
                if (playerPos[0] < 8) {
                    newPos = [playerPos[0] + 1, playerPos[1]];
                    //playerPos = [1, 4];
                    doorVal = 2;
                }
                break;
            case 'l':
                if (playerPos[1] > 0) {
                    newPos = [playerPos[0], playerPos[1] - 1];
                    //playerPos = [4, 1];
                    doorVal = 3;
                }
                break;
            case 'r':
                if (playerPos[1] < 8) {
                    newPos = [playerPos[0], playerPos[1] + 1];
                    //playerPos = [4, 7];
                    doorVal = 1;
                }
                break;
            default:
                break;
        }
        Excel.run(function (ctx) {
            var action = -1;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRange("b2:j10");
            var props = cellRange.getCellProperties({
                format: {
                    fill: {
                        color: true
                    }
                }
            });

            return ctx.sync()
                .then(function () {
                    if (props.value[newPos[0]][newPos[1]].format.fill.color == "#D3D3D3" || props.value[newPos[0]][newPos[1]].format.fill.color == "#FAEBD7" || props.value[newPos[0]][newPos[1]].format.fill.color == "#1E90FF") {
                        action = 0;
                    } else if (props.value[newPos[0]][newPos[1]].format.fill.color == "#FFFF00") {
                        action = 1;
                    }

                    switch (action) {
                        case 0:
                            movePlayerIcon(newPos[0], newPos[1]);
                            break;
                        case 1:
                            showNotification("Error", action + "");
                            moveRoom(roomList[currentRoom.doors[doorVal]]);
                            switch (doorVal) {
                                case 0:
                                    playerPos = [7, 4];
                                    break;
                                case 1:
                                    playerPos = [4, 1];
                                    break;
                                case 2:
                                    playerPos = [1, 4];
                                    break;
                                case 3:
                                    playerPos = [4, 7];
                                    break;
                            }
                            break;
                        default:
                            showNotification("Error", newPos[0] + " " + newPos[1] + "\n" + action);
                            break;
                    }
                })
                .then(ctx.sync);
        });
    }

    //Helper function for moving the player icon and updating the player position
    function movePlayerIcon(newRow, newCol) {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRange("b2:j10");

            return ctx.sync().then(function () {
                cellRange.getCell(playerPos[0], playerPos[1]).values = '';
                playerPos[0] = newRow;
                playerPos[1] = newCol;
                cellRange.getCell(playerPos[0], playerPos[1]).values = '☺';
            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    // Save this for the "Ron's Coming!" button
    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    //Function used to create all room tiles, their titles, descriptions, and item lists. Adds the rooms to the roomList
    //Called when application/game restarts
    function makeRooms() {

        roomList = [];

        /*
         *  0 = Chasm tile
         *  1 = Floor tile
         *  2 = Sand tile
         *  3 = Water tile
         *  4 = Wall tile
         *  5 = Door tile
         *  6 = Main Exit Tile
         */

        let startTiles = [
            [4, 4, 4, 6, 6, 6, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let startRoomItems = [
            {
                "item": '☻',
                "row": 1,
                "col": 1
            },
            {
                "item": '☻',
                "row": 1,
                "col": 5
            },
            {
                "item": '☻',
                "row": 3,
                "col": 1
            },
            {
                "item": '☻',
                "row": 3,
                "col": 5
            }
        ];
        let startRoomTitle = "Lobby of Freedom";
        let startRoomDesc = "Three keys bar this door,\nThree keys, and not one more\nI challenge thee to find them,\nI, the great Wizard weNnoR!";
        let startRoom = new room(startTiles, [16, -1, 1, -1], startRoomTitle, startRoomDesc, startRoomItems);
        roomList.push(startRoom);

        let roomTwoFourTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 0, 0, 1, 1, 1, 0, 0, 4],
            [4, 0, 0, 1, 0, 1, 0, 0, 4],
            [4, 0, 0, 1, 0, 1, 0, 0, 4],
            [4, 0, 0, 1, 0, 1, 0, 0, 4],
            [4, 0, 0, 1, 1, 1, 0, 0, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let twoFourItems = [];
        let twoFourTitle = "Salvation Bridge";
        let twoFourDesc = "One way leads to terror.\nOne way leads to freedom!";
        let roomTwoFour = new room(roomTwoFourTiles, [0, -1, 5, -1], twoFourTitle, twoFourDesc, twoFourItems);
        roomList.push(roomTwoFour);

        let lonelyIslandTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [4, 3, 3, 2, 2, 2, 3, 3, 4],
            [4, 3, 3, 2, 2, 2, 3, 2, 5],
            [4, 3, 3, 2, 2, 2, 3, 3, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let lonelyIslandItems = [
            {
                "item": 'F',
                "row": 3,
                "col": 3
            }
        ];
        let lonelyIslandTitle = "Lo'n Lee Island";
        let lonelyIslandDesc = "Save the key from its lonely perch";
        let lonelyIslandRoom = new room(lonelyIslandTiles, [-1, 3, -1, -1], lonelyIslandTitle, lonelyIslandDesc, lonelyIslandItems);
        roomList.push(lonelyIslandRoom);

        let tRoomTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 0, 0, 0, 0, 0, 0, 0, 4],
            [4, 0, 0, 0, 0, 0, 0, 0, 4],
            [4, 0, 0, 0, 0, 0, 0, 0, 4],
            [5, 2, 2, 1, 1, 1, 1, 1, 5],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 0, 0, 0, 1, 0, 0, 0, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let tRoomItems = [];
        let tRoomTitle = "Intersection of M'ist Urtee";
        let tRoomDesc = "I pity the fool that goes the wrong way!";
        let tRoom = new room(tRoomTiles, [-1, 4, 6, 2], tRoomTitle, tRoomDesc, tRoomItems);
        roomList.push(tRoom);

        let windyRoomTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 1, 1, 1, 3, 1, 1, 1, 4],
            [4, 1, 3, 1, 3, 1, 3, 1, 4],
            [4, 1, 3, 1, 3, 1, 3, 1, 4],
            [5, 1, 3, 1, 3, 1, 3, 1, 5],
            [4, 1, 3, 1, 3, 1, 3, 1, 4],
            [4, 1, 3, 1, 3, 1, 3, 1, 4],
            [4, 3, 3, 1, 1, 1, 3, 3, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let windyRoomItems = [];
        let windyRoomTitle = "Windig Zimmer";
        let windyRoomDesc = "Diese Zimmer ist sehr kurvenrich";
        let windyRoom = new room(windyRoomTiles, [-1, 5, 7, 3], windyRoomTitle, windyRoomDesc, windyRoomItems);
        roomList.push(windyRoom);

        let lungRoomTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1 ,1, 4],
            [4, 1, 1, 0, 0, 0, 1, 1, 4],
            [4, 1, 1, 0, 0, 0, 1, 1, 4],
            [5, 1, 1, 0, 1, 1, 1, 1, 4],
            [4, 1, 1, 0, 0, 0, 1, 1, 4],
            [4, 1, 1, 0, 0, 0, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];

        let lungRoomItems = [];
        let lungRoomTitle = "Breathing Room";
        let lungRoomDesc = "";
        let lungRoom = new room(lungRoomTiles, [1, -1, 8, 4], lungRoomTitle, lungRoomDesc, lungRoomItems);
        roomList.push(lungRoom);

        let bedRoomTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 4, 1, 4, 1, 1, 4],
            [4, 4, 4, 4, 1, 4, 4, 4, 4],
            [4, 4, 4, 4, 1, 4, 4, 4, 4],
            [4, 4, 4, 4, 1, 4, 4, 4, 4],
            [4, 1, 1, 4, 1, 4, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let bedRoomItems = [
            {
                "item": '∩',
                "row": 0,
                "col": 0
            },
            {
                "item": '∩',
                "row": 0,
                "col": 6
            },
            {
                "item": '∩',
                "row": 6,
                "col": 0
            },
            {
                "item": '∩',
                "row": 6,
                "col": 6
            }
        ];
        let bedRoomTitle = ""; //Need a name
        let bedRoomDesc = "";
        let bedRoom = new room(bedRoomTiles, [3, -1, 12, -1], bedRoomTitle, bedRoomDesc, bedRoomItems);
        roomList.push(bedRoom);

        let wennorTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let wennorItems = [
            {
                "item": 'i',
                "row": 0,
                "col": 0
            },
            {
                "item": 'i',
                "row": 0,
                "col": 6
            },
            {
                "item": '☺',
                "row": 1,
                "col": 1
            },
            {
                "item": '☺',
                "row": 1,
                "col": 3
            },
            {
                "item": '☺',
                "row": 1,
                "col": 5
            }, {
                "item": '╤',
                "row": 2,
                "col": 1
            },
            {
                "item": '╤',
                "row": 2,
                "col": 3
            },
            {
                "item": '╤',
                "row": 2,
                "col": 5
            },
            {
                "item": '☺',
                "row": 3,
                "col": 1
            },
            {
                "item": '☺',
                "row": 3,
                "col": 3
            },
            {
                "item": '☺',
                "row": 3,
                "col": 5
            }, {
                "item": '╤',
                "row": 4,
                "col": 1
            },
            {
                "item": '╤',
                "row": 4,
                "col": 3
            },
            {
                "item": '╤',
                "row": 4,
                "col": 5
            },
            {
                "item": 'W',
                "row": 6,
                "col": 3
            }
        ];
        let wennorTitle = "weNnoR's Auditorium";
        let wennorDesc = "Tsc fo daeh eht si weNnoR!";
        let wennorRoom = new room(wennorTiles, [4, -1, -1, -1], wennorTitle, wennorDesc, wennorItems);
        roomList.push(wennorRoom);

        let cornersTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 4, 4, 1, 4, 4, 1, 4],
            [4, 1, 4, 1, 1, 1, 4, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 5],
            [4, 1, 4, 1, 1, 1, 4, 1, 4],
            [4, 1, 4, 4, 1, 4, 4, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
        ];
        let cornersRoomItems = [
            {
                "items": 'i',
                "row": 2,
                "col": 2
            },
            {
                "items": 'i',
                "row": 2,
                "col": 4
            },
            {
                "items": 'i',
                "row": 4,
                "col": 2
            },
            {
                "items": 'i',
                "row": 4,
                "col": 4
            }
        ];
        let cornersRoomTitle = "Corners Room"; //MORE PUNS!
        let cornersRoomDesc = "";
        var cornersRoom = new room(cornersTiles, [5, 9, 14, -1], cornersRoomTitle, cornersRoomDesc, cornersRoomItems);
        roomList.push(cornersRoom);

        let lakeTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [5, 1, 1, 3, 3, 3, 3, 1, 5],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [4, 1, 1, 3, 3, 3, 3, 3, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
        ];
        let lakeItems = [];
        let lakeTitle = "Lac du Lac"; //Maybe keep this one?
        let lakeDesc = "Moisture is the essence of wetness.";
        let lakeRoom = new room(lakeTiles, [-1, 10, -1, 8], lakeTitle, lakeDesc, lakeItems);
        roomList.push(lakeRoom);

        let roundRoomTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 0, 0, 0, 0, 0, 0, 0, 4],
            [4, 0, 0, 1, 1, 1, 0, 0, 4],
            [4, 0, 1, 1, 0, 1, 1, 0, 4],
            [5, 1, 1, 0, 0, 0, 1, 1, 4],
            [4, 0, 1, 1, 0, 1, 1, 0, 4],
            [4, 0, 0, 1, 1, 1, 0, 0, 4],
            [4, 0, 0, 0, 0, 0, 0, 0, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
        ];
        let roundItems = [];
        let roundTitle = "E Pluribus Anus"; //Temp Name
        let roundDesc = "";
        let roundRoom = new room(roundRoomTiles, [-1, -1, -1, 9], roundTitle, roundDesc, roundItems);
        roomList.push(roundRoom);

        let commaRoomTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 0, 0, 0, 0, 0, 0, 1, 4],
            [4, 0, 0, 0, 1, 0, 0, 1, 4],
            [4, 1, 1, 0, 1, 0, 0, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 5],
            [4, 1, 1, 0, 1, 0, 0, 1, 4],
            [4, 0, 0, 0, 1, 0, 0, 1, 4],
            [4, 0, 0, 0, 0, 0, 0, 1, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let commaItems = [
            {
                "item": '/',
                "row": 1,
                "col": 3
            },
            {
                "item": '≡',
                "row": 3,
                "col": 0
            },
            {
                "item": '/',
                "row": 5,
                "col": 3
            }
        ];
        let commaTitle = "Tomb of O'xf-Ord";
        let commaDesc = "";
        let commaRoom = new room(commaRoomTiles, [-1, 12, -1, -1], commaTitle, commaDesc, commaItems);
        roomList.push(commaRoom);

        let skullTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 0, 0, 1, 0, 0, 1, 4],
            [4, 1, 0, 0, 1, 0, 0, 1, 4],
            [5, 1, 0, 0, 1, 0, 0, 1, 5],
            [4, 1, 1, 0, 1, 0, 1, 1, 4],
            [4, 0, 1, 1, 1, 1, 1, 0, 4],
            [4, 0, 1, 0, 1, 0, 1, 0, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let skullItems = [];
        let skullTitle = "Kingdom of the Crystal Skull"; //Temp name
        let skullDesc = "";
        let skullRoom = new room(skullTiles, [6, 13, -1, 11], skullTitle, skullDesc, skullItems);
        roomList.push(skullRoom);

        let bunnyEarTiles = [
            [4, 4, 4, 4, 4, 4, 4, 4, 4],
            [4, 1, 1, 1, 0, 1, 1, 1, 4],
            [4, 1, 0, 1, 0, 1, 0, 1, 4],
            [4, 1, 0, 1, 0, 1, 0, 1, 4],
            [5, 1, 0, 1, 0, 1, 0, 1, 5],
            [4, 1, 0, 1, 0, 1, 0, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 0, 1, 1, 1, 1, 1, 0, 4],
            [4, 4, 4, 4, 5, 4, 4, 4, 4]
        ];
        let bunnyItems = [];
        let bunnyTitle = "Hoppy Brewery"; //Temp name
        let bunnyDesc = "";
        let bunnyRoom = new room(bunnyEarTiles, [-1, 14, 15, 12], bunnyTitle, bunnyDesc, bunnyItems);
        roomList.push(bunnyRoom);

        let torchIslandTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 3, 3, 1, 1, 1, 3, 3, 4],
            [4, 3, 3, 1, 1, 1, 3, 3, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [5, 1, 1, 3, 3, 3, 1, 1, 4],
            [4, 3, 3, 3, 3, 3, 3, 3, 4],
            [4, 3, 3, 1, 1, 1, 3, 3, 4],
            [4, 3, 3, 1, 1, 1, 3, 3, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let torchIslandItems = [
            {
                "item": 'i',
                "row": 6,
                "col": 3
            }
        ];
        let torchIslandTitle = "Torch Island";
        let torchIslandDesc = "";
        let torchIslandRoom = new room(torchIslandTiles, [8, -1, -1, 13], torchIslandTitle, torchIslandDesc, torchIslandItems);
        roomList.push(torchIslandRoom);

        let snorkelTiles = [
            [4, 4, 4, 4, 5, 4, 4, 4, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 1, 3, 3, 3, 3, 3, 1, 4],
            [4, 1, 3, 3, 3, 3, 3, 1, 4],
            [4, 1, 3, 3, 3, 3, 3, 1, 4],
            [4, 1, 3, 3, 3, 3, 3, 1, 4],
            [4, 1, 1, 1, 1, 1, 1, 1, 4],
            [4, 0, 0, 1, 1, 1, 0, 0, 4],
            [4, 4, 4, 4, 4, 4, 4, 4, 4]
        ];
        let snorkelItems = [
            {
                "item": 'J',
                "row": 6,
                "col": 3
            }
        ];
        let snorkelTitle = "Sanctuary of Snor'Kel";
        let snorkelDesc = "";
        let snorkelRoom = new room(snorkelTiles, [13, -1, -1, -1], snorkelTitle, snorkelDesc, snorkelItems);
        roomList.push(snorkelRoom);

        return roomList;
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
