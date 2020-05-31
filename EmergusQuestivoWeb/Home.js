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

    //let startTiles = [
    //    ['f', 'f', 'f', 'c', 'f', 'f', 'k'],
    //    ['f', 'c', 'f', 'c', 'f', 'c', 'f'],
    //    ['f', 'c', 'f', 'c', 'f', 'c', 'f'],
    //    ['f', 'c', 'f', 'c', 'f', 'c', 'f'],
    //    ['f', 'c', 'p', 'c', 'f', 'c', 'f'],
    //    ['w', 'f', 'f', 'f', 'f', 'f', 'w'],
    //    ['w', 'w', 'w', 'w', 'w', 'w', 'w']
    //];

    //var startRoom = new room(startTiles, [true, true, false, true], "Start Room");

    //Tracking the player icon's position
    var playerPos = [3, 4];
    var roomList = [];

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

            $("#template-description").text("Navigate your way through the Wizard weNnoR's realm. Find three keys (ᚩ)");
            $('#button-text').text("Show me Potato Salad!");
            $('#button-desc').text("Highlights the largest number.");

            $('#btn-up-text').text("Up");
            $('#btn-up').click(function () {
                move('u');
            });

            $('#btn-down-text').text("Down");
            $('#btn-down').click(function () {
                move('d');
            });

            $('#btn-left-text').text("Left");
            $('#btn-left').click(function () {
                move('l');
            });

            $('#btn-right-text').text("Right");
            $('#btn-right').click(function () {
                move('r');
            });

            //setCellSizes();
            roomList = makeRooms();
            //loadSampleData();
            moveRoom(roomList[0]);
            //Set player location
            
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

            return ctx.sync();
            //return context.sync()
            //    .then(function () {
            //        if (sheets.items.length === 1) {
            //            console.log("Can't delete the last sheet");
            //        } else {
            //            currentSheet.delete();
            //        }
            //    });
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

    //Render rooms with a 2 cell pad on top and left sides (top left room edge starts at Row 3, Column C)
    function roomRender(newRoom) {
        //newRoom = roomList[0];

        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            //var wallRange = sheet.getRange("b2:j10");
            //wallRange.load("value, rowCount, columnCount");
            //c3:i9

            var cellRange = sheet.getRange("b2:j10");
            cellRange.load("value, rowCount, columnCount");

            return ctx.sync().then(function () {
                //wallRange.format.fill.color = "SaddleBrown"; //#8B4513

                //if (newRoom.doors[0] >= 0) {
                //    wallRange.getCell(0, 4).format.fill.color = "Orange"; //#FFA500
                //}
                //if (newRoom.doors[1] >= 0) {
                //    wallRange.getCell(4, 8).format.fill.color = "Orange";
                //}
                //if (newRoom.doors[2] >= 0) {
                //    wallRange.getCell(8, 4).format.fill.color = "Orange";
                //}
                //if (newRoom.doors[3] >= 0) {
                //    wallRange.getCell(4, 0).format.fill.color = "Orange";
                //}

                for (var i = 0; i < cellRange.rowCount; i++) {
                    for (var j = 0; j < cellRange.columnCount; j++) {
                        switch (newRoom.tiles[i][j]) {
                            case 0:
                                cellRange.getCell(i, j).format.fill.color = "Black";
                                break;
                            case 1:
                                cellRange.getCell(i, j).format.fill.color = "LightGrey";
                                break;
                            case 2:
                                cellRange.getCell(i, j).fromat.fill.color = "AntiqueWhite";
                                break;
                            case 3:
                                cellRange.getCell(i, j).format.fill.color = "DodgerBlue";
                                break;
                            case 4:
                                cellRange.getCell(i, j).format.fill.color = "SaddleBrown"; //#8B4513
                                break;
                            case 5:
                                cellRange.getCell(i, j).format.fill.color = "Yellow"; //#FFFF00
                                break;
                            case 'k':
                                cellRange.getCell(i, j).format.fill.color = "brown";
                                cellRange.getCell(i, j).values = 'ᚩ';
                                break;
                            case 'p':
                                cellRange.getCell(i, j).format.fill.color = "brown";
                                cellRange.getCell(i, j).values = '☺';

                                //Update player's position
                                playerPos[0] = i;
                                playerPos[1] = j;
                                //showNotification("Position: " + playerPos[0] + ", " + playerPos[1]);
                                break;
                            default:
                                cellRange.getCell(i, j).format.fill.color = "black";
                                break;
                        }
                    }

                    //Display the player
                    cellRange.getCell(playerPos[0], playerPos[1]).values = '☺';
                }
            }).then(ctx.sync);
                
        }).catch(errorHandler);
    }

    //Function for checking a cell in a given direction, 
    //and if it's possible to move to the cell, 
    //move the player icon
    //and update player position
    function move(direction, currentRoom) {
        //showNotification("Dir: " + direction +  " Position: " + playerPos[0] + ", " + playerPos[1]);
        //Switch for directions
        switch (direction) {
            case 'u':
                if (playerPos[0] > 0) {
                    //checkValidMove(playerPos[0] - 1, playerPos[1]);
                    switch (checkValidMove(playerPos[0] - 1, playerPos[1])) {
                        case 0:
                            movePlayerIcon(playerPos[0] - 1, playerPos[1]);
                            break;
                        case 1:
                            moveRoom(roomList[currentRoom.doors[0]]);
                            break;
                        default:
                            break;
                    }
                }
                break;
            case 'd':
                if (playerPos[0] < 8) {
                    switch (checkValidMove(playerPos[0] + 1, playerPos[1])) {
                        case 0:
                            movePlayerIcon(playerPos[0] + 1, playerPos[1]);
                            break;
                        case 1:
                            moveRoom(roomList[currentRoom.doors[2]]);
                            break;
                        default:
                            break;
                    }
                }
                break;
            case 'l':
                if (playerPos[1] > 0) {
                    switch (checkValidMove(playerPos[0], playerPos[1] - 1)) {
                        case 0:
                            movePlayerIcon(playerPos[0], playerPos[1] - 1);
                            break;
                        case 1:
                            moveRoom(roomList[currentRoom.doors[3]]);
                            break;
                        default:
                            break;
                    }
                }
                break;
            case 'r':
                if (playerPos[1] < 8) {
                    switch (checkValidMove(playerPos[0], playerPos[1] + 1)) {
                        case 0:
                            movePlayerIcon(playerPos[0], playerPos[1] + 1);
                            break;
                        case 1:
                            moveRoom(roomList[currentRoom.doors[1]]);
                            break;
                        default:
                            break;
                    }
                }
                break;
            default:
                break;
        }
    }

    //Helper function for checking if a given cell is valid for movement
    function checkValidMove(row, col) {
        var rVal = -1;

        Excel.run(function (ctx) {
            //var rVal = -1;
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var cellRange = sheet.getRange("b2:j10");
            //cellRange.load("format/fill/color");
            var props = cellRange.getCellProperties({
                format: {
                    fill: {
                        color: true
                    },
                },
            });
            return ctx.sync().then(function () {
                //console.log(props.value);
                console.log(props.value[row][col].format.fill.color == "#A52A2A");
                if (props.value[row][col].format.fill.color == "#D3D3D3") {
                    //movePlayerIcon(row, col);
                    rVal = 0;
                } else if (props.value[row][col].format.fill.color == "#FFFF00") {
                    rVal = 1;
                }

                return rVal;
            }).then(ctx.sync);
        }).catch(errorHandler);

        return rVal;
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

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function makeRooms() {

        roomList = [];

        /*
         *  0 = Chasm tile
         *  1 = Floor tile
         *  2 = Sand tile
         *  3 = Water tile
         *  4 = Wall tile
         *  5 = Door tile
         */

        let startTiles = [
            [4, 4, 4, 5, 5, 5, 4, 4, 4],
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
                "row": 1,
                "col": 5
            }
        ];
        let startRoomTitle = "Lobby of Freedom";
        let startRoomDesc = "Three keys bar this door,\nThree keys, and not one more\nI challenge thee to find them,\nI, the great Wizard weNnoR!";
        let startRoom = new room(startTiles, [16, -1, 2, -1], startRoomTitle, startRoomDesc, startRoomItems);
        roomList.push(startRoom);

        let roomTwoFourTiles = [
            [0, 0, 0, 1, 0, 0, 0],
            [0, 0, 1, 1, 1, 0, 0],
            [0, 0, 1, 0, 1, 0, 0],
            [0, 0, 1, 0, 1, 0, 0],
            [0, 0, 1, 0, 1, 0, 0],
            [0, 0, 1, 1, 1, 0, 0],
            [0, 0, 0, 1, 0, 0, 0]
        ];
        let twoFourItems = [];
        let twoFourTitle = "Salvation Bridge";
        let twoFourDesc = "One way leads to terror.\nOne way leads to freedom!";
        let roomTwoFour = new room(roomTwoFourTiles, [0, -1, 5, -1], twoFourTitle, twoFourDesc, twoFourItems);
        roomList.push(roomTwoFour);

        let lonelyIslandTiles = [
            [3, 3, 3, 3, 3, 3, 3],
            [3, 3, 3, 3, 3, 3, 3],
            [3, 3, 2, 2, 2, 3, 3],
            [3, 3, 2, 2, 2, 3, 2],
            [3, 3, 2, 2, 2, 3, 3],
            [3, 3, 3, 3, 3, 3, 3],
            [3, 3, 3, 3, 3, 3, 3],
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
            [0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0],
            [0, 0, 0, 0, 0, 0, 0],
            [2, 2, 1, 1, 1, 1, 1],
            [0, 0, 0, 1, 0, 0, 0],
            [0, 0, 0, 1, 0, 0, 0],
            [0, 0, 0, 1, 0, 0, 0]
        ];
        let tRoomItems = [];
        let tRoomTitle = "Intersection of M'ist Urtee";
        let tRoomDesc = "I pity the fool that goes the wrong way!";
        let tRoom = new room(tRoomTiles, [-1, 4, 6, 2], tRoomTitle, tRoomDesc, tRoomItems);
        roomList.push(tRoom);

        let windyRoomTiles = [
            [1, 1, 1, 3, 1, 1, 1],
            [1, 3, 1, 3, 1, 3, 1],
            [1, 3, 1, 3, 1, 3, 1],
            [1, 3, 1, 3, 1, 3, 1],
            [1, 3, 1, 3, 1, 3, 1],
            [1, 3, 1, 3, 1, 3, 1],
            [3, 3, 1, 1, 1, 3, 3]
        ];
        let windyRoomItems = [];
        let windyRoomTitle = "Windig Zimmer";
        let windyRoomDesc = "Diese Zimmer ist sehr kurvenrich";
        let windyRoom = new room(windyRoomTiles, [-1, 5, -1, 3], windyRoomTitle, windyRoomDesc, windyRoomItems);
        roomList.push(windyRoom);

        let lungRoomTiles = [
            [1, 1, 1, 1, 1, 1 ,1],
            [1, 1, 0, 0, 0, 1, 1],
            [1, 1, 0, 0, 0, 1, 1],
            [1, 1, 0, 1, 1, 1, 1],
            [1, 1, 0, 0, 0, 1, 1],
            [1, 1, 0, 0, 0, 1, 1],
            [1, 1, 1, 1, 1, 1, 1]
        ];

        let lungRoomItems = [];
        let lungRoomTitle = "Breathing Room";
        let lungRoomDesc = "";
        let lungRoom = new room(lungRoomTiles, [1, -1, 8, 4], lungRoomTitle, lungRoomDesc, lungRoomItems);
        roomList.push(lungRoom);

        let bedRoomTiles = [
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 4, 1, 4, 1, 1],
            [4, 4, 4, 1, 4, 4, 4],
            [4, 4, 4, 1, 4, 4, 4],
            [4, 4, 4, 1, 4, 4, 4],
            [1, 1, 4, 1, 4, 1, 1],
            [1, 1, 1, 1, 1, 1, 1]
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
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 1, 1, 1, 1, 1],
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
            [1, 1, 1, 1, 1, 1, 1],
            [1, 4, 4, 1, 4, 4, 1],
            [1, 4, 1, 1, 1, 4, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 4, 1, 1, 1, 4, 1],
            [1, 4, 4, 1, 4, 4, 1],
            [1, 1, 1, 1, 1, 1, 1]
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
            [1, 1, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 3, 1],
            [1, 1, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 3, 3]
        ];
        let lakeItems = [];
        let lakeTitle = "Lac du Lac"; //Maybe keep this one?
        let lakeDesc = "Moisture is the essence of wetness.";
        let lakeRoom = new room(lakeTiles, [-1, 10, -1, 8], lakeTitle, lakeDesc, lakeItems);
        roomList.push(lakeRoom);

        let roundRoomTiles = [
            [0, 0, 0, 0, 0, 0, 0],
            [0, 0, 1, 1, 1, 0, 0],
            [0, 1, 1, 0, 1, 1, 0],
            [1, 1, 0, 0, 0, 1, 1],
            [0, 1, 1, 0, 1, 1, 0],
            [0, 0, 1, 1, 1, 0, 0],
            [0, 0, 0, 0, 0, 0, 0]
        ];
        let roundItems = [];
        let roundTitle = "E Pluribus Anus"; //Temp Name
        let roundDesc = "";
        let roundRoom = new room(roundRoomTiles, [-1, -1, -1, 9], roundTitle, roundDesc, roundItems);
        roomList.push(roundRoom);

        let commaRoomTiles = [
            [0, 0, 0, 0, 0, 0, 1],
            [0, 0, 0, 1, 0, 0, 1],
            [1, 1, 0, 1, 0, 0, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [1, 1, 0, 1, 0, 0, 1],
            [0, 0, 0, 1, 0, 0, 1],
            [0, 0, 0, 0, 0, 0, 1]
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
            [1, 1, 1, 1, 1, 1, 1],
            [1, 0, 0, 1, 0, 0, 1],
            [1, 0, 0, 1, 0, 0, 1],
            [1, 0, 0, 1, 0, 0, 1],
            [1, 1, 0, 1, 0, 1, 1],
            [0, 1, 1, 1, 1, 1, 0],
            [0, 1, 0, 1, 0, 1, 0]
        ];
        let skullItems = [];
        let skullTitle = "Kingdom of the Crystal Skull"; //Temp name
        let skullDesc = "";
        let skullRoom = new room(skullTiles, [6, 13, -1, 11], skullTitle, skullDesc, skullItems);

        let bunnyEarTiles = [
            [1, 1, 1, 0, 1, 1, 1],
            [1, 0, 1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1, 0, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [0, 1, 1, 1, 1, 1, 0]
        ];
        let bunnyItems = [];
        let bunnyTitle = "Hoppy Brewery"; //Temp name
        let bunnyDesc = "";
        let bunnyRoom = new room(bunnyEarTiles, [-1, 14, 15, 12], bunnyTitle, bunnyDesc, bunnyItems);
        roomList.push(bunnyRoom);

        let torchIslandTiles = [
            [3, 3, 1, 1, 1, 3, 3],
            [3, 3, 1, 1, 1, 3, 3],
            [3, 3, 3, 3, 3, 3, 3],
            [1, 1, 3, 3, 3, 1, 1],
            [3, 3, 3, 3, 3, 3, 3],
            [3, 3, 1, 1, 1, 3, 3],
            [3, 3, 1, 1, 1, 3, 3]
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
        let torchIslandRoom = new room(torchIslandTiles, [-1, 14, 15, 12], torchIslandTitle, torchIslandDesc, torchIslandItems);
        roomList.push(torchIslandRoom);

        let snorkelTiles = [
            [1, 1, 1, 1, 1, 1, 1],
            [1, 3, 3, 3, 3, 3, 1],
            [1, 3, 3, 3, 3, 3, 1],
            [1, 3, 3, 3, 3, 3, 1],
            [1, 3, 3, 3, 3, 3, 1],
            [1, 1, 1, 1, 1, 1, 1],
            [0, 0, 1, 1, 1, 0, 0]
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
