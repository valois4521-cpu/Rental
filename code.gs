// ----------------------------------------------------------------
// 定数設定
// ----------------------------------------------------------------
const SPREADSHEET_ID = '1hEqJrDvqy-7p--ZHEhGd8Khw2fh6PvC7GqeWC_W3KFU';
// ★検査成績表の保存先フォルダID (ユーザー指定)
const INSPECTION_FOLDER_ID = '1a8Ut1oDKvcKu3PG95zIJ-r686_sfQRRI';

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const itemSheet = ss.getSheetByName('物品マスタ');
const rentalSheet = ss.getSheetByName('貸出履歴');
const mailSheet = ss.getSheetByName('メールマスタ'); // メール通知用のシート
const categorySheet = ss.getSheetByName('カテゴリマスタ'); // ★追加

// 物品マスタのヘッダー定義 
const HEADER = ['管理番号', '物品名', 'カテゴリ', 'ステータス', '貸与者', '現場名', '貸与日', '返却予定日', '校正日・入替日', '備考', '検査成績表'];
// 貸出履歴シートのヘッダー定義
const HISTORY_HEADER = ['履歴ID', '管理番号', '物品名', '処理', '貸与者', '現場名', '貸与日', '返却予定日', '実返却日', '通知先アドレス'];
const JST = "Asia/Tokyo"

// ----------------------------------------------------------------
// Webページ表示
// ----------------------------------------------------------------
function doGet(e) {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('貸与品管理アプリ')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ----------------------------------------------------------------
// データ取得
// ----------------------------------------------------------------

/**
 * 初期表示に必要な全データを取得する
 */
function getInitialData() {
    try {
        const itemValues = itemSheet && itemSheet.getLastRow() > 1 ? itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, HEADER.length).getValues() : [];
        const rentalValues = rentalSheet && rentalSheet.getLastRow() > 1 ? rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, HISTORY_HEADER.length).getValues() : [];

        try {
            const categoryOrder = getCategoryOrderMap();
            
            if (categoryOrder.size > 0) {
                itemValues.sort((a, b) => {
                    const catIndex = HEADER.indexOf('カテゴリ');
                    const catA = String(a[catIndex] || '').trim();
                    const catB = String(b[catIndex] || '').trim();
                    
                    const orderA = categoryOrder.has(catA) ? categoryOrder.get(catA) : 9999;
                    const orderB = categoryOrder.has(catB) ? categoryOrder.get(catB) : 9999;

                    if (orderA !== orderB) {
                        return orderA - orderB;
                    }
                    const idA = a[0] || '';
                    const idB = b[0] || '';
                    return idA.localeCompare(idB, undefined, { numeric: true, sensitivity: 'base' });
                });
            }
        } catch (sortError) {
            console.warn("並び替え処理中にエラーが発生しましたが、無視して続行します: " + sortError.message);
        }

        const items = itemValues.map(row => formatRowAsObject(row, HEADER));
        const rentals = rentalValues.map(row => formatRowAsObject(row, HISTORY_HEADER));

        return { items: items, rentals: rentals };
    } catch(e) {
        Logger.log("getInitialData Error: " + e.message + e.stack);
        return { items: [], rentals: [] };
    }
}

/**
 * カテゴリマスタからカテゴリ名と並び順のマッピングを作成する
 */
function getCategoryOrderMap() {
    const map = new Map();
    if (!categorySheet) return map;

    try {
        const lastRow = categorySheet.getLastRow();
        if (lastRow < 2) return map;

        const values = categorySheet.getRange(2, 1, lastRow - 1, 2).getValues();

        values.forEach((row, index) => {
            const name = String(row[0] || '').trim();
            const order = row[1];
            if (name) {
                const sortKey = (typeof order === 'number') ? order : index + 1;
                map.set(name, sortKey);
            }
        });
    } catch (e) {
        console.warn("カテゴリマスタ読み込みエラー: " + e.message);
    }
    return map;
}

/**
 * 物品の貸出制約（次の予約）と今後のスケジュールを取得する
 */
function getItemScheduleInfo(itemId) {
    const futureEvents = getFutureEventsForItem(itemId);
    const nextReservation = futureEvents.find(event => event[3] === '予約');
    
    let constraints = {
        latestReturnDate: null, 
        message: ''
    };

    if (nextReservation) {
        const nextReservationStartDate = normalizeDate(nextReservation[6]);
        const latestReturnDate = new Date(nextReservationStartDate.getTime());
        latestReturnDate.setDate(latestReturnDate.getDate() - 1); 
        constraints.latestReturnDate = Utilities.formatDate(latestReturnDate, JST, 'yyyy-MM-dd');
        constraints.message = `注意: この物品は${Utilities.formatDate(nextReservationStartDate, JST, 'M月d日')}から予約が入っています。`;
    }

    const schedule = futureEvents.map(event => {
        return {
            type: event[3], 
            borrower: event[4],
            siteName: event[5] || '',
            startDate: Utilities.formatDate(normalizeDate(event[6]), JST, 'yyyy/MM/dd'),
            endDate: Utilities.formatDate(normalizeDate(event[7]), JST, 'yyyy/MM/dd'),
        }
    });

    return { constraints: constraints, schedule: schedule };
}


/**
 * 特定の物品の編集可能なスケジュール（現在の貸出と今後の予約）を取得する
 */
function getEditableSchedulesForItem(itemId) {
    const lastRow = rentalSheet.getLastRow();
    if (lastRow < 2) return [];
    const values = rentalSheet.getRange(2, 1, lastRow - 1, HISTORY_HEADER.length).getValues();
    const today = normalizeDate(new Date());

    const schedules = values
        .filter(row => {
            const isTargetItem = row[1] == itemId;
            const isEditableStatus = ['貸出', '予約'].includes(row[3]);
            const eventEndDate = normalizeDate(row[7]);
            return isTargetItem && isEditableStatus && eventEndDate >= today;
        })
        .map(row => {
            const history = formatRowAsObject(row, HISTORY_HEADER);
            return {
                historyId: history.履歴ID,
                type: history.処理,
                borrower: history.貸与者,
                siteName: history.現場名,
                startDate: history.貸与日,
                returnDate: history.返却予定日,
                itemName: history.物品名,
                email: history.通知先アドレス || '' 
            };
        })
        .sort((a,b) => normalizeDate(a[6]) - normalizeDate(b[6])); 
    
    return schedules;
}


// ----------------------------------------------------------------
// データ操作
// ----------------------------------------------------------------

/**
 * 新規物品を登録する
 */
function addItem(itemData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        if (!itemData.管理番号) {
            throw new Error('管理番号が指定されていません。');
        }
        const existingRow = findRowById(itemSheet, itemData.管理番号);
        if (existingRow !== -1) {
            throw new Error('その管理番号は既に使用されています。');
        }
        
        itemData.ステータス = '在庫';
        itemData.貸与者 = '';
        itemData.貸与日 = ''; 
        itemData.現場名 = '';
        itemData.返却予定日 = '';

        const newRow = HEADER.map(key => {
            if (['校正日・入替日'].includes(key) && itemData[key]) {
                return normalizeDate(itemData[key]); 
            }
            return itemData[key] || '';
        });
        itemSheet.appendRow(newRow);
        
        return { success: true, message: '新しい物品を登録しました。', shouldReload: true };
    } catch (e) { 
        Logger.log('Error in addItem: ' + e.message + e.stack);
        return { success: false, message: '登録に失敗しました: ' + e.message }; 
    } finally {
        lock.releaseLock();
    }
}

/**
 * 既存の物品情報を更新する
 */
function updateItem(itemData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const originalId = itemData.original管理番号; 
        const newId = itemData.管理番号; 

        const itemRow = findRowById(itemSheet, originalId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');
        
        if (originalId !== newId) {
            const existingRow = findRowById(itemSheet, newId);
            if (existingRow !== -1) {
                throw new Error(`管理番号「${newId}」は既に使用されています。`);
            }
        }
        
        const itemRange = itemSheet.getRange(itemRow, 1, 1, HEADER.length);
        const currentValues = itemRange.getValues()[0];
        const newRowData = HEADER.map((key, index) => {
            if (itemData.hasOwnProperty(key)) {
                if (['校正日・入替日'].includes(key) && itemData[key]) {
                    return normalizeDate(itemData[key]); 
                }
                return itemData[key] || '';
            }
            return currentValues[index]; 
        });
        itemRange.setValues([newRowData]);

        if (originalId !== newId) {
            const rentalValues = rentalSheet.getDataRange().getValues();
            const idColIndex = HISTORY_HEADER.indexOf('管理番号');
            
            for (let i = 1; i < rentalValues.length; i++) { 
                if (String(rentalValues[i][idColIndex]) === String(originalId)) {
                    rentalSheet.getRange(i + 1, idColIndex + 1).setValue(newId);
                }
            }
        }
        
        return { success: true, message: '物品情報を更新しました。', shouldReload: true };
    } catch (e) { 
        Logger.log('Error in updateItem: ' + e.message + e.stack);
        return { success: false, message: '更新に失敗しました: ' + e.message }; 
    } finally {
        lock.releaseLock();
    }
}

/**
 * 物品を削除する
 */
function deleteItem(itemId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        const status = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
        if(status === '貸出中') throw new Error('貸出中の物品は削除できません。');

        const futureEvents = getFutureEventsForItem(itemId);
        if (futureEvents.length > 0) throw new Error('この物品には予約が入っているため削除できません。');

        itemSheet.deleteRow(itemRow);
        return { success: true, message: '物品を削除しました。', shouldReload: true };
    } catch (e) { 
        Logger.log('Error in deleteItem: ' + e.message + e.stack);
        return { success: false, message: `削除に失敗しました: ${e.message}` }; 
    } finally {
        lock.releaseLock();
    }
}

/**
 * 検査成績表をアップロードし、物品マスタにURLを保存する
 * @param {object} data { itemId, fileName, mimeType, data (base64) }
 */
function uploadInspectionReport(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, fileName, mimeType, data: base64Data } = data;
        const decoded = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decoded, mimeType, `検査成績表_${itemId}_${fileName}`);

        let folder;
        try {
            folder = DriveApp.getFolderById(INSPECTION_FOLDER_ID);
        } catch (e) {
            Logger.log('フォルダアクセスエラー: ' + e.message);
            throw new Error(
                '検査成績表の保存先フォルダにアクセスできません。\n' +
                '【対処法】GASエディタで「doAuthorize」関数を手動実行し、Googleドライブの権限を承認してから、新しいバージョンで再デプロイしてください。\n' +
                '(フォルダID: ' + INSPECTION_FOLDER_ID + ')'
            );
        }

        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = file.getUrl();

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        const colIndex = HEADER.indexOf('検査成績表');
        if (colIndex === -1) throw new Error('物品マスタに「検査成績表」列がありません。');

        itemSheet.getRange(itemRow, colIndex + 1).setValue(fileUrl);

        return { success: true, message: '検査成績表を保存しました。' };
    } catch (e) {
        Logger.log('Error in uploadInspectionReport: ' + e.message + e.stack);
        return { success: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

/**
 * 検査成績表を削除する（Driveのファイルをゴミ箱に移動し、物品マスタのURLをクリア）
 * @param {string} itemId 管理番号
 */
function deleteInspectionReport(itemId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        const colIndex = HEADER.indexOf('検査成績表');
        if (colIndex === -1) throw new Error('物品マスタに「検査成績表」列がありません。');

        const fileUrl = itemSheet.getRange(itemRow, colIndex + 1).getValue();
        if (!fileUrl) throw new Error('検査成績表が登録されていません。');

        // Google DriveのURLからファイルIDを抽出してゴミ箱へ移動
        try {
            const fileIdMatch = fileUrl.match(/[-\w]{25,}/);
            if (fileIdMatch) {
                const file = DriveApp.getFileById(fileIdMatch[0]);
                file.setTrashed(true);
            }
        } catch (e) {
            Logger.log('Driveファイル削除時の警告: ' + e.message);
            // ファイルが既に削除済みでもURL欄のクリアは続行する
        }

        // 物品マスタの検査成績表URLをクリア
        itemSheet.getRange(itemRow, colIndex + 1).setValue('');

        return { success: true, message: '検査成績表を削除しました。', shouldReload: true };
    } catch (e) {
        Logger.log('Error in deleteInspectionReport: ' + e.message + e.stack);
        return { success: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

// ----------------------------------------------------------------
// 貸出・返却・予約 処理
// ----------------------------------------------------------------

/**
 * 物品を貸し出す
 */
function lendItem(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, startDate, returnDate, email } = data;
        if (!borrower) throw new Error('貸与者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();

        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            const today = normalizeDate(new Date());
            if (normalizeDate(maintenanceDate) < today) {
                throw new Error(`この物品は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を過ぎているため、貸し出せません。`);
            }
        }
        
        const currentStatus = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
        if (['貸出中'].includes(currentStatus)) {
            throw new Error('この物品は現在貸出中のため、新たに貸し出すことはできません。');
        }
        
        const checkResult = checkAvailability(itemId, new Date(startDate), new Date(returnDate));
        if (!checkResult.isAvailable) {
            throw new Error(`貸出期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        addHistoryRecord(itemId, null, '貸出', data);
        updateItemStatus(itemId);

        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) { 
            const itemName = getItemNameById(itemId);
            const subject = `【貸与品管理】物品貸出通知 (${itemName})`;
            const body =
`${borrower} 様（または現場担当者 様）

以下の内容で物品の貸出が実行されました。

物品名: ${itemName} (管理番号: ${itemId})
貸与者: ${borrower}
現場名: ${siteName || 'N/A'}
貸出日: ${Utilities.formatDate(normalizeDate(startDate), JST, 'yyyy/MM/dd')}
返却予定日: ${Utilities.formatDate(normalizeDate(returnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '編集', itemCategory);
        }

        return { success: true, message: '貸出処理が完了しました。', shouldReload: true };
    } catch (e) { 
        Logger.log('Error in lendItem: ' + e.message + e.stack);
        return { success: false, message: `貸出処理に失敗しました: ${e.message}` }; 
    } finally {
        lock.releaseLock();
    }
}

/**
 * 物品を返却する
 */
function returnItem(historyId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const historyRow = findRowById(rentalSheet, historyId);
        if (historyRow === -1) throw new Error('対象の履歴が見つかりません。');

        const historyRange = rentalSheet.getRange(historyRow, 1, 1, HISTORY_HEADER.length);
        const historyValues = historyRange.getValues()[0];
        
        const currentStatus = historyValues[HISTORY_HEADER.indexOf('処理')];
        if (currentStatus === '返却済み') {
            throw new Error('この貸出は既に返却済みです。');
        }

        historyValues[HISTORY_HEADER.indexOf('処理')] = '返却済み';
        historyValues[HISTORY_HEADER.indexOf('実返却日')] = new Date();
        historyRange.setValues([historyValues]);

        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];
        updateItemStatus(itemId);

        const itemCategory = getItemCategoryById(itemId);

        const email = historyValues[HISTORY_HEADER.indexOf('通知先アドレス')];
        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const borrower = historyValues[HISTORY_HEADER.indexOf('貸与者')];
            const siteName = historyValues[HISTORY_HEADER.indexOf('現場名')];
            const startDate = historyValues[HISTORY_HEADER.indexOf('貸与日')];
            const returnDate = historyValues[HISTORY_HEADER.indexOf('実返却日')]; 

            const subject = `【貸与品管理】物品返却通知 (${itemName})`;
            const body =
`${borrower} 様（または現場担当者 様）

以下の物品が返却されました。

物品名: ${itemName} (管理番号: ${itemId})
貸与者: ${borrower}
現場名: ${siteName || 'N/A'}
貸出日: ${Utilities.formatDate(normalizeDate(startDate), JST, 'yyyy/MM/dd')}
実返却日: ${Utilities.formatDate(normalizeDate(returnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '返却', itemCategory);
        }

        return { success: true, message: '返却処理が完了しました。', shouldReload: true };
    } catch (e) { 
        Logger.log('Error in returnItem: ' + e.message + e.stack);
        return { success: false, message: '返却処理に失敗しました: ' + e.message }; 
    } finally {
        lock.releaseLock();
    }
}

/**
 * 貸出中の物品の貸与者、現場名、期間、通知先アドレスを更新する
 */
function updateRentalPeriod(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, newStartDate, newReturnDate, email } = data;
        if (!borrower) throw new Error('貸与者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();

        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            if (normalizeDate(newReturnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`返却日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
        }

        const historyValues = rentalSheet.getDataRange().getValues();
        if (historyValues.length < 2) throw new Error('更新対象の貸出履歴が見つかりません。');
        
        const historyRowIndex = historyValues.length - 1 - [...historyValues].reverse().findIndex(row => row[1] == itemId && row[3] === '貸出');
        if (historyRowIndex < 1) throw new Error('更新対象の貸出履歴が見つかりません。');
        
        const historyIdToIgnore = historyValues[historyRowIndex][0];
        const checkResult = checkAvailability(itemId, new Date(newStartDate), new Date(newReturnDate), true, historyIdToIgnore);
        if (!checkResult.isAvailable) {
            throw new Error(`変更後の期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        itemSheet.getRange(itemRow, HEADER.indexOf('貸与者') + 1).setValue(borrower);
        itemSheet.getRange(itemRow, HEADER.indexOf('現場名') + 1).setValue(siteName);
        itemSheet.getRange(itemRow, HEADER.indexOf('貸与日') + 1).setValue(Utilities.formatDate(normalizeDate(newStartDate), JST, 'yyyy/MM/dd'));
        itemSheet.getRange(itemRow, HEADER.indexOf('返却予定日') + 1).setValue(Utilities.formatDate(normalizeDate(newReturnDate), JST, 'yyyy/MM/dd'));
        
        const historyRow = historyRowIndex + 1;
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('貸与者') + 1).setValue(borrower);
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('現場名') + 1).setValue(siteName);
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('貸与日') + 1).setValue(normalizeDate(newStartDate));
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('返却予定日') + 1).setValue(normalizeDate(newReturnDate));
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('通知先アドレス') + 1).setValue(email || '');
        
        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const subject = `【貸与品管理】貸出内容変更通知 (${itemName})`;
            const body =
`貸出内容が変更されました。

物品名: ${itemName} (管理番号: ${itemId})
貸与者: ${borrower}
現場名: ${siteName || 'N/A'}
新しい貸出日: ${Utilities.formatDate(normalizeDate(newStartDate), JST, 'yyyy/MM/dd')}
新しい返却予定日: ${Utilities.formatDate(normalizeDate(newReturnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '編集', itemCategory);
        }

        return { success: true, message: '貸出内容を更新しました。', shouldReload: true };
    } catch (e) {
        Logger.log('Error in updateRentalPeriod: ' + e.message + e.stack);
        return { success: false, message: `貸出内容の更新に失敗しました: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * 物品を予約する
 */
function addReservation(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, startDate, returnDate, email } = data;
        if (!borrower) throw new Error('予約者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');
        
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();

        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            const today = normalizeDate(new Date());
            if (normalizeDate(returnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`予約終了日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
            if (normalizeDate(maintenanceDate) < today) {
                throw new Error(`この物品は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を過ぎているため、予約できません。`);
            }
        }

        const checkResult = checkAvailability(itemId, new Date(startDate), new Date(returnDate));
        if (!checkResult.isAvailable) {
            throw new Error(`予約期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        addHistoryRecord(itemId, null, '予約', data);
        updateItemStatus(itemId);

        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const subject = `【貸与品管理】物品予約通知 (${itemName})`;
            const body =
`${borrower} 様（または現場担当者 様）

以下の内容で物品の予約が実行されました。

物品名: ${itemName} (管理番号: ${itemId})
予約者: ${borrower}
現場名: ${siteName || 'N/A'}
予約開始日: ${Utilities.formatDate(normalizeDate(startDate), JST, 'yyyy/MM/dd')}
予約終了日: ${Utilities.formatDate(normalizeDate(returnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '予約', itemCategory);
        }

        return { success: true, message: '予約を受け付けました。', shouldReload: true };
    } catch (e) {
        Logger.log('Error in addReservation: ' + e.message + e.stack);
        return { success: false, message: `予約に失敗しました: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * 予約内容を更新する
 */
function updateReservation(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { historyId, borrower, siteName, newStartDate, newReturnDate, email } = data;
        if (!borrower) throw new Error('予約者が入力されていません。');
        
        const historyRow = findRowById(rentalSheet, historyId);
        if (historyRow === -1) throw new Error('対象の予約が見つかりません。');

        const historyRange = rentalSheet.getRange(historyRow, 1, 1, HISTORY_HEADER.length);
        const historyValues = historyRange.getValues()[0];
        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。'); 
        
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();

        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            if (normalizeDate(newReturnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`予約終了日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
        }

        const checkResult = checkAvailability(itemId, new Date(newStartDate), new Date(newReturnDate), false, historyId);
        if (!checkResult.isAvailable) {
            throw new Error(`変更後の期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        historyValues[HISTORY_HEADER.indexOf('貸与者')] = borrower;
        historyValues[HISTORY_HEADER.indexOf('現場名')] = siteName;
        historyValues[HISTORY_HEADER.indexOf('貸与日')] = normalizeDate(newStartDate);
        historyValues[HISTORY_HEADER.indexOf('返却予定日')] = normalizeDate(newReturnDate);
        historyValues[HISTORY_HEADER.indexOf('通知先アドレス')] = email || '';
        
        historyRange.setValues([historyValues]);
        
        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const subject = `【貸与品管理】予約内容変更通知 (${itemName})`;
            const body =
`予約内容が変更されました。

物品名: ${itemName} (管理番号: ${itemId})
予約者: ${borrower}
現場名: ${siteName || 'N/A'}
新しい予約開始日: ${Utilities.formatDate(normalizeDate(newStartDate), JST, 'yyyy/MM/dd')}
新しい予約終了日: ${Utilities.formatDate(normalizeDate(newReturnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '編集', itemCategory);
        }
        
        return { success: true, message: '予約内容を更新しました。', shouldReload: true };
    } catch (e) {
        Logger.log('Error in updateReservation: ' + e.message + e.stack);
        return { success: false, message: `予約の更新に失敗しました: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * 予約をキャンセルする
 */
function cancelReservation(historyId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const historyRow = findRowById(rentalSheet, historyId);
        if (historyRow === -1) throw new Error('対象の予約が見つかりません。');
        
        const historyRange = rentalSheet.getRange(historyRow, 1, 1, HISTORY_HEADER.length);
        const historyValues = historyRange.getValues()[0];
        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];
        const email = historyValues[HISTORY_HEADER.indexOf('通知先アドレス')];

        historyValues[HISTORY_HEADER.indexOf('処理')] = 'キャンセル済';
        historyRange.setValues([historyValues]);
        
        updateItemStatus(itemId);
        
        const itemCategory = getItemCategoryById(itemId);

        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const borrower = historyValues[HISTORY_HEADER.indexOf('貸与者')];
            const siteName = historyValues[HISTORY_HEADER.indexOf('現場名')];
            const startDate = historyValues[HISTORY_HEADER.indexOf('貸与日')];
            const returnDate = historyValues[HISTORY_HEADER.indexOf('返却予定日')];

            const subject = `【貸与品管理】予約キャンセル通知 (${itemName})`;
            const body =
`${borrower} 様（または現場担当者 様）

以下の予約がキャンセルされました。

物品名: ${itemName} (管理番号: ${itemId})
予約者: ${borrower}
現場名: ${siteName || 'N/A'}
予約開始日: ${Utilities.formatDate(normalizeDate(startDate), JST, 'yyyy/MM/dd')}
予約終了日: ${Utilities.formatDate(normalizeDate(returnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationEmail(email, subject, body, '予約', itemCategory);
        }

        return { success: true, message: '予約をキャンセルしました。', shouldReload: true };
    } catch(e) {
        Logger.log('Error in cancelReservation: ' + e.message + e.stack);
        return { success: false, message: `予約のキャンセルに失敗しました: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * 予約を「貸出」ステータスに変更する
 */
function lendFromReservation(historyId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const historyRow = findRowById(rentalSheet, historyId);
        if (historyRow === -1) throw new Error('対象の予約が見つかりません。');

        const historyRange = rentalSheet.getRange(historyRow, 1, 1, HISTORY_HEADER.length);
        const historyValues = historyRange.getValues()[0];
        
        const currentStatus = historyValues[HISTORY_HEADER.indexOf('処理')];
        if (currentStatus !== '予約') {
            throw new Error('この履歴は予約ではないため、貸出処理を実行できません。');
        }
        const startDate = normalizeDate(historyValues[HISTORY_HEADER.indexOf('貸与日')]);
        const today = normalizeDate(new Date());

        if (startDate > today) {
            throw new Error(`この予約は${Utilities.formatDate(startDate, JST, 'M月d日')}開始のため、まだ貸し出せません。`);
        }

        historyValues[HISTORY_HEADER.indexOf('処理')] = '貸出';
        historyRange.setValues([historyValues]);
        
        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];
        const email = historyValues[HISTORY_HEADER.indexOf('通知先アドレス')];
        updateItemStatus(itemId); 
        
        const itemCategory = getItemCategoryById(itemId);

        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const borrower = historyValues[HISTORY_HEADER.indexOf('貸与者')];
            const siteName = historyValues[HISTORY_HEADER.indexOf('現場名')];
            const startDate = historyValues[HISTORY_HEADER.indexOf('貸与日')];
            const returnDate = historyValues[HISTORY_HEADER.indexOf('返却予定日')];

            const subject = `【貸与品管理】貸出開始通知 (${itemName})`;
            const body =
`${borrower} 様（または現場担当者 様）

以下の予約が貸出中に変更されました。

物品名: ${itemName} (管理番号: ${itemId})
貸与者: ${borrower}
現場名: ${siteName || 'N/A'}
貸出日: ${Utilities.formatDate(normalizeDate(startDate), JST, 'yyyy/MM/dd')}
返却予定日: ${Utilities.formatDate(normalizeDate(returnDate), JST, 'yyyy/MM/dd')}

---
貸与品管理アプリ`;

            sendNotificationNotificationEmail(email, subject, body, '編集', itemCategory);
        }

        return { success: true, message: '予約を貸出中に変更しました。', shouldReload: true };
    } catch(e) {
        Logger.log('Error in lendFromReservation: ' + e.message + e.stack);
        return { success: false, message: `貸出処理に失敗しました: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

// ----------------------------------------------------------------
// カレンダー用データ取得関数
// ----------------------------------------------------------------

/**
 * FullCalendarの「リソース」（物品一覧）データを取得する
 */
function getItemsAsResources() {
    try {
        const itemValues = itemSheet.getLastRow() > 1 ? itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, HEADER.length).getValues() : [];
        return itemValues.map(row => ({ 
            id: row[0], // 管理番号
            title: row[1], // 物品名
            category: row[2] // カテゴリ (グループ化用)
        }));
    } catch (e) { 
        Logger.log('Error in getItemsAsResources: ' + e.message + e.stack);
        return []; 
    }
}

/**
 * FullCalendarの「イベント」（貸出・予約・校正日）データを取得する
 */
function getCalendarEvents() {
    const rentalLastRow = rentalSheet.getLastRow();
    const rentalValues = rentalLastRow < 2 ? [] : rentalSheet.getRange(2, 1, rentalLastRow - 1, HISTORY_HEADER.length).getValues();

    const rentalEvents = rentalValues
        .filter(row => ['貸出', '予約', '返却済み'].includes(row[3])) 
        .map(row => {
            const history = formatRowAsObject(row, HISTORY_HEADER);
            const startDate = normalizeDate(history.貸与日);
            
            let effectiveEndDate = normalizeDate(history.返却予定日);
            if (history.処理 === '返却済み' && history.実返却日) {
                effectiveEndDate = normalizeDate(history.実返却日);
            }

            const endForCalendar = new Date(effectiveEndDate.getTime());
            endForCalendar.setDate(endForCalendar.getDate() + 1);

            let eventData = {
                id: history.履歴ID,
                resourceId: history.管理番号,
                start: Utilities.formatDate(startDate, JST, 'yyyy-MM-dd'),
                end: Utilities.formatDate(endForCalendar, JST, 'yyyy-MM-dd'),
                extendedProps: { 
                    historyId: history.履歴ID,
                    itemId: history.管理番号, 
                    itemName: history.物品名, 
                    borrower: history.貸与者,
                    startDate: history.貸与日, 
                    returnDate: history.返却予定日,
                    actualReturnDate: history.実返却日 || null,
                    siteName: history.現場名,
                    email: history.通知先アドレス || ''
                }
            };

            switch(history.処理) {
                case '貸出':
                    eventData.title = `${history.現場名 || history.貸与者} (貸出)`;
                    eventData.backgroundColor = '#d9534f'; 
                    eventData.borderColor = '#d43f3a';
                    eventData.extendedProps.type = 'rental';
                    break;
                case '予約':
                    eventData.title = `${history.現場名 || history.貸与者} (予約)`;
                    eventData.backgroundColor = '#0d6efd'; 
                    eventData.borderColor = '#0b5ed7';
                    eventData.extendedProps.type = 'reservation';
                    break;
                case '返却済み':
                    eventData.title = `【返却済み】${history.現場名 || history.貸与者}`;
                    eventData.backgroundColor = '#adb5bd'; 
                    eventData.borderColor = '#6c757d';
                    eventData.extendedProps.type = 'returned';
                    break;
            }
            return eventData;
        });

    const itemLastRow = itemSheet.getLastRow();
    const itemValues = itemLastRow < 2 ? [] : itemSheet.getRange(2, 1, itemLastRow - 1, HEADER.length).getValues();
    
    const maintenanceEvents = itemValues
        .map(row => formatRowAsObject(row, HEADER))
        .filter(item => item['校正日・入替日']) 
        .map(item => {
            const maintenanceDate = normalizeDate(item['校正日・入替日']);
            return {
                id: `maint_${item.管理番号}`,
                resourceId: item.管理番号,
                title: `校正/入替: ${item.物品名}`,
                start: Utilities.formatDate(maintenanceDate, JST, 'yyyy-MM-dd'),
                allDay: true,
                backgroundColor: '#6c757d', 
                borderColor: '#6c757d',
                extendedProps: {
                    type: 'maintenance',
                    itemId: item.管理番号,
                    itemName: item.物品名,
                    maintenanceDate: Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')
                }
            };
        });
    
    return [...rentalEvents, ...maintenanceEvents]; 
}


// ----------------------------------------------------------------
// スケジュール設定（トリガー用）
// ----------------------------------------------------------------

/**
 * 予約開始日当日の予約を自動で「貸出」ステータスに変更する（トリガーで毎日実行）
 */
function processReservations() {
    Logger.log("予約の自動貸出処理を開始します。");
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(300000)) { 
        Logger.log("他のプロセスが実行中のため、スキップします。");
        return;
    }

    try {
        const today = normalizeDate(new Date());
        const todayStr = Utilities.formatDate(today, JST, 'yyyy/MM/dd');

        if (rentalSheet.getLastRow() < 2) {
            Logger.log("処理対象の予約はありません。");
            return;
        }
        const events = rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, HISTORY_HEADER.length).getValues();

        events.forEach((event, index) => {
            const status = event[3];
            const startDate = normalizeDate(event[6]);
            const startDateStr = Utilities.formatDate(startDate, JST, 'yyyy/MM/dd');
            
            if (status === '予約' && startDateStr === todayStr) {
                const itemId = event[1];
                const itemRow = findRowById(itemSheet, itemId);
                
                if (itemRow !== -1) {
                    const itemStatus = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
                    if (itemStatus === '予約あり') {
                        Logger.log(`予約ID[${event[0]}]を自動で貸し出します。`);
                        rentalSheet.getRange(index + 2, HISTORY_HEADER.indexOf('処理') + 1).setValue('貸出');
                        SpreadsheetApp.flush(); 
                        updateItemStatus(itemId); 
                    } else {
                        Logger.log(`物品[${itemId}]が返却されていないため、予約ID[${event[0]}]の自動貸出をスキップしました。`);
                    }
                }
            }
        });
        Logger.log("予約の自動貸出処理が完了しました。");
    } catch (e) {
        Logger.log("予約の自動貸出処理中にエラーが発生しました: ", e.message + e.stack);
    } finally {
        lock.releaseLock();
    }
}

// ----------------------------------------------------------------
// ヘルパー関数
// ----------------------------------------------------------------

/**
 * 貸出履歴シートに新しいレコードを追加する
 */
function addHistoryRecord(itemId, itemName, action, data) {
    const id = getNextId(rentalSheet); 
    const name = itemName || (itemId ? getItemNameById(itemId) : '不明な物品'); 
    const newRow = [
        id, itemId, name, action, data.borrower,
        data.siteName || '',
        normalizeDate(data.startDate), normalizeDate(data.returnDate),
        '', 
        data.email || '' 
    ];
    rentalSheet.appendRow(newRow);
}

/**
 * 物品マスタのステータスと貸出情報を最新の状態に更新する
 */
function updateItemStatus(itemId) {
    const itemRow = findRowById(itemSheet, itemId);
    if (itemRow === -1) return;

    const itemRange = itemSheet.getRange(itemRow, 1, 1, HEADER.length);
    const itemValues = itemRange.getValues()[0];
    const allFutureEvents = getFutureEventsForItem(itemId);
    const today = normalizeDate(new Date());
    
    const activeRental = allFutureEvents.find(e => e[3] === '貸出');

    if (activeRental) {
        itemValues[HEADER.indexOf('ステータス')] = '貸出中';
        itemValues[HEADER.indexOf('貸与者')] = activeRental[4];
        itemValues[HEADER.indexOf('現場名')] = activeRental[5];
        
        const lendDate = activeRental[6];
        const formattedLendDate = Utilities.formatDate(lendDate, JST, 'yyyy/MM/dd');
        
        itemValues[HEADER.indexOf('貸与日')] = formattedLendDate;
        itemValues[HEADER.indexOf('返却予定日')] = Utilities.formatDate(activeRental[7], JST, 'yyyy/MM/dd');

    } else {
        const upcomingReservation = allFutureEvents.find(e => 
            e[3] === '予約' && normalizeDate(e[6]) >= today
        );
        
        itemValues[HEADER.indexOf('ステータス')] = upcomingReservation ? '予約あり' : '在庫';
        itemValues[HEADER.indexOf('貸与者')] = '';
        itemValues[HEADER.indexOf('現場名')] = '';
        itemValues[HEADER.indexOf('貸与日')] = '';
        itemValues[HEADER.indexOf('返却予定日')] = '';
    }
    itemRange.setValues([itemValues]); 
}

/**
 * 特定の物品の、今日以降に終了する「貸出」または「予約」の履歴を取得する
 */
function getFutureEventsForItem(itemId) {
    const lastRow = rentalSheet.getLastRow();
    if (lastRow < 2) return [];
    const values = rentalSheet.getRange(2, 1, lastRow - 1, HISTORY_HEADER.length).getValues();
    const today = normalizeDate(new Date());

    return values
        .filter(row => {
            const isTargetItem = row[1] == itemId;
            const isFutureEvent = ['貸出', '予約'].includes(row[3]);
            const eventEndDate = normalizeDate(row[7]);
            return isTargetItem && isFutureEvent && eventEndDate >= today;
        })
        .sort((a,b) => normalizeDate(a[6]) - normalizeDate(b[6])); 
}

/**
 * 指定された期間が、他の既存の予定と重複していないかチェックする
 */
function checkAvailability(itemId, newStart, newEnd, isEditingLend = false, historyIdToIgnore = null) {
    const normalizedStart = normalizeDate(newStart);
    const normalizedEnd = normalizeDate(newEnd);

    const futureEvents = getFutureEventsForItem(itemId);

    for (const event of futureEvents) {
        if (String(event[0]) === String(historyIdToIgnore)) continue;

        const eventStart = normalizeDate(event[6]);
        const eventEnd = normalizeDate(event[7]);
        
        if (normalizedStart <= eventEnd && normalizedEnd >= eventStart) {
            const reason = `${event[3]}あり (${Utilities.formatDate(eventStart, JST, 'yyyy/MM/dd')} ~ ${Utilities.formatDate(eventEnd, JST, 'yyyy/MM/dd')})`;
            return { isAvailable: false, conflictReason: reason };
        }
    }
    return { isAvailable: true }; 
}

/**
 * 指定されたシートのA列からIDを検索し、その行番号（1始まり）を返す
 */
function findRowById(sheet, id) {
    if (!id || sheet.getLastRow() < 2) return -1;
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(cellId => String(cellId) == String(id)); 
    return (rowIndex !== -1) ? rowIndex + 2 : -1; 
}

/**
 * 物品IDから物品のカテゴリを取得する
 */
function getItemCategoryById(itemId) {
    try {
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) return '';
        return itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();
    } catch (e) {
        Logger.log(`getItemCategoryById Error: ${e.message}`);
        return '';
    }
}

/**
 * スプレッドシートの行データ（配列）を、ヘッダーに基づいたオブジェクトに変換する
 */
function formatRowAsObject(rowArray, headerArray) {
    let item = {};
    const DATE_KEYS = ['貸出日', '返却予定日', '校正日・入替日', '貸与日', '実返却日']; 
    headerArray.forEach((key, index) => {
        const value = rowArray[index];
        
        if (DATE_KEYS.includes(key)) {
            const normalized = normalizeDate(value);
            if (normalized) {
                item[key] = Utilities.formatDate(normalized, JST, 'yyyy/MM/dd');
            } else {
                item[key] = ''; 
            }
        } else {
            item[key] = value;
        }
    });
    return item;
}

/**
 * 貸出履歴シートのA列の最大ID+1の値を返す
 */
function getNextId(sheet) {
    const lastRow = rentalSheet.getLastRow();
    if (lastRow < 2) return 1; 
    const ids = rentalSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const maxId = Math.max(0, ...ids.filter(id => !isNaN(id)).map(id => Number(id)));
    return isFinite(maxId) ? maxId + 1 : 1;
}

/**
 * タイムゾーンのズレを吸収し、JSTの0時0分0秒のDateオブジェクトを返す
 */
function normalizeDate(date) {
    if (!date) return null;

    try {
        let d;
        if (date instanceof Date) {
            const dateString = Utilities.formatDate(date, JST, 'yyyy/MM/dd');
            d = new Date(dateString); 
        } 
        else if (typeof date === 'string') {
            const dateString = date.split('T')[0].replace(/-/g, '/');
            d = new Date(dateString); 
        } 
        else {
            const tempDate = new Date(date);
            const dateString = Utilities.formatDate(tempDate, JST, 'yyyy/MM/dd');
            d = new Date(dateString);
        }
        
        if (isNaN(d.getTime())) { 
            return null;
        }

        return d; 
    } catch(e) {
        Logger.log("normalizeDate Error: " + e.message + " (Input: " + date + ")");
        return null;
    }
}

// ----------------------------------------------------------------
// メール送信用ヘルパー関数
// ----------------------------------------------------------------

/**
 * メールマスタシートから指定された種類かつカテゴリに対応するメールアドレスを取得する
 */
function getMailMasterAddresses(mailType, targetCategory) {
    if (typeof mailSheet === 'undefined' || !mailSheet) {
        Logger.log("「メールマスタ」シートが見つかりません。");
        return [];
    }
    try {
        const lastRow = mailSheet.getLastRow();
        if (lastRow < 2) return [];

        const values = mailSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        
        const candidates = values.filter(row => row[1] === mailType && row[2]);

        const targetCat = targetCategory ? String(targetCategory).trim() : '';

        const specificMatches = candidates.filter(row => {
            const rowCategory = String(row[3] || '').trim();
            if (!rowCategory) return false; 

            const categories = rowCategory.split(',').map(c => c.trim());
            return categories.includes(targetCat);
        });

        let finalRows = [];

        if (specificMatches.length > 0) {
            finalRows = specificMatches;
        } else {
            finalRows = candidates.filter(row => !String(row[3] || '').trim());
        }

        const addresses = finalRows.map(row => row[2].trim())
            .filter(address => address.length > 0 && address.includes('@')); 

        return [...new Set(addresses)]; 
    } catch (e) {
        Logger.log(`getMailMasterAddresses Error: ${e.message} ${e.stack}`);
        return [];
    }
}

/**
 * 操作完了通知メールを送信する
 */
function sendNotificationEmail(formEmail, subject, body, mailType, targetCategory) {
    let recipients = [];

    if (formEmail) {
        const emailsFromForm = formEmail.split(',')
            .map(email => email.trim())
            .filter(email => email.length > 0 && email.includes('@'));
        recipients = recipients.concat(emailsFromForm);
    }

    const type = mailType || '編集';
    const masterAddresses = getMailMasterAddresses(type, targetCategory);
    recipients = recipients.concat(masterAddresses);

    const uniqueRecipients = [...new Set(recipients)].filter(email => email);

    if (uniqueRecipients.length > 0) {
        try {
            const recipientString = uniqueRecipients.join(',');
            MailApp.sendEmail({
                to: recipientString,
                subject: subject,
                body: body,
            });
            Logger.log(`メール送信成功: ${recipientString} (Type: ${type}, Category: ${targetCategory})`);
        } catch (e) {
            Logger.log(`メール送信エラー: ${e.message} ${e.stack}`);
        }
    } else {
        Logger.log(`通知先アドレスが指定されていないため、メールは送信されませんでした。(Type: ${type}, Category: ${targetCategory})`);
    }
}

/**
 * 物品IDから物品名を取得する
 */
function getItemNameById(itemId) {
    try {
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) return `不明な物品(ID: ${itemId})`;
        return itemSheet.getRange(itemRow, HEADER.indexOf('物品名') + 1).getValue();
    } catch (e) {
        Logger.log(`getItemNameById Error: ${e.message}`);
        return `不明な物品(ID: ${itemId})`;
    }
}

/**
 * 【重要】権限承認用関数
 * この関数はアプリのロジックでは使用しません。
 * GASエディタで手動実行し、Googleドライブの読み書き権限を承認してください。
 *
 * 手順:
 *   1. GASエディタでこの関数を選択して「実行」をクリック
 *   2. 権限の承認ダイアログが表示されたら「許可」する
 *   3. 実行ログで「成功」と表示されることを確認
 *   4. デプロイ → デプロイを管理 → 新しいバージョンで再デプロイ
 */
function doAuthorize() {
  // フォルダへの読み取りアクセス確認
  const folder = DriveApp.getFolderById(INSPECTION_FOLDER_ID);
  console.log("フォルダ「" + folder.getName() + "」への読み取りOK");

  // ファイル書き込み権限の確認（テストファイルを作成→即削除）
  const testFile = folder.createFile('_権限テスト.txt', 'テスト', 'text/plain');
  testFile.setTrashed(true);
  console.log("成功: ファイル書き込み権限もOKです。新しいバージョンで再デプロイしてください。");
}
