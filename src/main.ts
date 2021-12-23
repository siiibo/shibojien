import { BlockAction } from '@slack/bolt'
import { SlashCommand } from '@slack/bolt'
import { RespondArguments } from '@slack/bolt';
import { Button, MessageAttachment } from '@slack/types';

const NAME_COL = 9;
const INDEX_COL = 10;
const EXPLANATION_COL = 11;

function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
    slackHandler(e);
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
}

function isBlockAction(e: GoogleAppsScript.Events.DoPost): boolean {
    if (e.parameter.hasOwnProperty('payload')) {
        const type = JSON.parse(e.parameter['payload'])['type'];
        return type === 'block_actions';
    }
    return false;
}

function isSlachCommand(e: GoogleAppsScript.Events.DoPost): boolean {
    return e.parameter.hasOwnProperty('command');
}

function slackHandler(e: GoogleAppsScript.Events.DoPost){
    if (isBlockAction(e)) {
        responseKiyopediaAction(JSON.parse(e.parameter['payload']));
    }
    else if (isSlachCommand(e)) {
        receiveKiyopediaCommand(e.parameter);
    }
    else {
        return;
    }
}

function receiveKiyopediaCommand(parameter: SlashCommand){
    const query = parameter.text;
    if (query == "") {
        const response: RespondArguments = {
            response_type: 'ephemeral',
            text: 'コマンドの後は *探したい言葉* 又は *言葉の一部* を続けてください（コマンドの直後にはスペースが必要です）。\n'
                 + '例：/jien しぼさい、/jien 私募債',
        }       
        respond(parameter.response_url, response)
        return;
    }

    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHIBOJIEN_SPREADSHEET_ID')).getSheets()[0];
    const words = sheet.getRange(1, NAME_COL, sheet.getLastRow()).getValues();

    const cells = sheet.createTextFinder(query).findAll().filter((range)=>{
        const col = range.getColumn();
        return col == INDEX_COL || col == NAME_COL;
    });

    let preRow = -1;
    let foundWords = [];
    for (const cell of cells) {
        const row = cell.getRow();
        if (preRow == row && row <= 2) continue;

        foundWords.push(words[row-1][0]);
        preRow = row;
    }

    if (foundWords.length <= 0) {
        // const url = PropertiesService.getScriptProperties().getProperty('SHIBOJIEN_SPREADSHEET_URL');
        const message = 'そのような読み仮名を持つ言葉は存在しませんでした。\n'
                        + '※ひらがなの方がヒットしやすいですが、漢字・カタカナ・アルファベットでも検索可能です。';
        const response: RespondArguments = {
            response_type: 'in_channel',
            text: message,
        }
        respond(parameter.response_url, response);
        return;
    }
    else {
        let elements = [];
        for (const word of foundWords) {
            const element: Button = {
                type: 'button',
                text: {
                    type: 'plain_text',
                    text: word,
                    emoji: true,
                },
                value: word,
            }
            elements.push(element);
        }

        const attachments: MessageAttachment = {
            color: '#00bfff',
            blocks: [
                {
                    type: 'actions',
                    elements: elements,
                }
            ]   
        }

        const response: RespondArguments = {
            response_type: 'in_channel',
            text: foundWords.length+'件ヒットしました。',
            attachments: [attachments],
        }
        respond(parameter.response_url, response);
        return;
    }
}


function responseKiyopediaAction(payload: BlockAction){
    const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHIBOJIEN_SPREADSHEET_ID')).getSheets()[0];

    const query = payload['actions'][0]['value'];
    const cells = sheet.getRange(1, NAME_COL, sheet.getLastRow()).createTextFinder(query)
        .matchEntireCell(true)
        .findAll();

    if (cells.length <= 0) {
        const message = '*参照エラー*';
        const response: RespondArguments = {
            response_type: 'ephemeral',
            text: message,
        }

        respond(payload.response_url, response);
        return;
    }
    else {
        const row = cells[0].getRow();
        const name = cells[0].getValue();
        const explanation = sheet.getRange(row, EXPLANATION_COL).getValue();
        
        const attachments: MessageAttachment = {
            color: '#00bfff',
            blocks: [
				{
					type: 'section',
					text: {
						type: 'mrkdwn',
						text: '*【'+name+'】*',
					}
				},
                {
                    type: 'section',
					text: {
						type: 'mrkdwn',
						text: explanation,
					}
                }
            ]
        }
        
        const response: RespondArguments = {
            response_type: 'in_channel',
            text: 'お探しの言葉の社内定義はこちらです。',
            attachments: [attachments],
        }
        respond(payload.response_url, response);
        return;
    }
}

function respond(response_url: string, payload: any) {
    const option = {
        method: 'post',
        muteHttpExceptions: true,
        validateHttpsCertificates: false,
        followRedirects: false,
        payload: JSON.stringify(payload),
    }
    const res = UrlFetchApp.fetch(response_url, option);
}

function init() {
    initProperties();
}

function initProperties() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROPERTIES');
    const rows = sheet.getDataRange().getValues();
    let properties = {};
    for (let row of rows.slice(1)) properties[row[0]] = row[1];

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
    scriptProperties.setProperties(properties);
}

declare const global: any;
global.doPost = doPost;
global.init = init;

