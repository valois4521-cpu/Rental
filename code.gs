// ----------------------------------------------------------------
// 定数設定
// ----------------------------------------------------------------
const SPREADSHEET_ID = '1gIbD9KboeQs-wsfWAn97S3NGNZlqDXUtFm29Rd6DUgE';
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const itemSheet = ss.getSheetByName('物品マスタ');
const rentalSheet = ss.getSheetByName('貸出履歴');
const mailSheet = ss.getSheetByName('メールマスタ'); // メール通知用のシート

// 物品マスタのヘッダー定義 
const HEADER = ['管理番号', '物品名', 'カテゴリ', 'ステータス', '貸与者', '現場名', '貸与日', '返却予定日', '校正日・入替日', '備考'];
// 貸出履歴シートのヘッダー定義
const HISTORY_HEADER = ['履歴ID', '管理番号', '物品名', '処理', '貸与者', '現場名', '貸与日', '返却予定日', '実返却日', '通知先アドレス'];
const JST = "Asia/Tokyo"

// ----------------------------------------------------------------
// Webページ表示
// ----------------------------------------------------------------
function doGet(e) {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Kounai_rental')
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
        const itemValues = itemSheet.getLastRow() > 1 ? itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, HEADER.length).getValues() : [];
        const rentalValues = rentalSheet.getLastRow() > 1 ? rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, HISTORY_HEADER.length).getValues() : [];

        // 物品マスタを管理番号順にソート
        itemValues.sort((a, b) => {
            const idA = a[0] || '';
            const idB = b[0] || '';
            return idA.localeCompare(idB, undefined, { numeric: true, sensitivity: 'base' });
        });

        const items = itemValues.map(row => formatRowAsObject(row, HEADER));
        const rentals = rentalValues.map(row => formatRowAsObject(row, HISTORY_HEADER));

        return { items: items, rentals: rentals };
    } catch(e) {
        Logger.log("getInitialData Error: " + e.message + e.stack);
        return { items: [], rentals: [] };
    }
}

/**
 * 物品の貸出制約（次の予約）と今後のスケジュールを取得する
 * @param {string} itemId 管理番号
 * @return {object} 制約情報とスケジュールリスト
 */
function getItemScheduleInfo(itemId) {
    const futureEvents = getFutureEventsForItem(itemId);
    const nextReservation = futureEvents.find(event => event[3] === '予約');
    
    let constraints = {
        latestReturnDate: null, // 貸出可能な最終返却日
        message: ''
    };

    // 次に予約がある場合、その前日までしか貸し出せないように制約を設ける
    if (nextReservation) {
        const nextReservationStartDate = normalizeDate(nextReservation[6]);
        const latestReturnDate = new Date(nextReservationStartDate.getTime());
        latestReturnDate.setDate(latestReturnDate.getDate() - 1); // 予約日の前日
        constraints.latestReturnDate = Utilities.formatDate(latestReturnDate, JST, 'yyyy-MM-dd');
        constraints.message = `注意: この物品は${Utilities.formatDate(nextReservationStartDate, JST, 'M月d日')}から予約が入っています。`;
    }

    // フロントエンド表示用にスケジュールを整形
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
 * @param {string} itemId 管理番号
 * @return {Array} スケジュールオブジェクトの配列
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
            // 返却予定日が今日以降のもの（＝まだ終わっていない予定）を対象
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
 * @param {object} itemData 物品データ
 */
function addItem(itemData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        if (!itemData.管理番号) {
            throw new Error('管理番号が指定されていません。');
        }
        // 管理番号の重複チェック
        const existingRow = findRowById(itemSheet, itemData.管理番号);
        if (existingRow !== -1) {
            throw new Error('その管理番号は既に使用されています。');
        }
        
        // 新規登録時はステータスを「在庫」に設定
        itemData.ステータス = '在庫';
        itemData.貸与者 = '';
        itemData.貸与日 = ''; // HEADER定義の「貸与日」キー
        itemData.現場名 = '';
        itemData.返却予定日 = '';

        const newRow = HEADER.map(key => {
            if (['校正日・入替日'].includes(key) && itemData[key]) {
                return normalizeDate(itemData[key]); // 日付を正規化
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
 * @param {object} itemData 物品データ (original管理番号 を含む)
 */
function updateItem(itemData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const originalId = itemData.original管理番号; // 変更前のID
        const newId = itemData.管理番号; // 変更後のID

        const itemRow = findRowById(itemSheet, originalId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');
        
        // IDが変更された場合、新しいIDが既に存在しないかチェック
        if (originalId !== newId) {
            const existingRow = findRowById(itemSheet, newId);
            if (existingRow !== -1) {
                throw new Error(`管理番号「${newId}」は既に使用されています。`);
            }
        }
        
        // 物品マスタの行データを更新
        const itemRange = itemSheet.getRange(itemRow, 1, 1, HEADER.length);
        const currentValues = itemRange.getValues()[0];
        const newRowData = HEADER.map((key, index) => {
            if (itemData.hasOwnProperty(key)) {
                if (['校正日・入替日'].includes(key) && itemData[key]) {
                    return normalizeDate(itemData[key]); // 日付を正規化
                }
                return itemData[key] || '';
            }
            return currentValues[index]; // 変更しない項目は元の値
        });
        itemRange.setValues([newRowData]);

        // IDが変更された場合、貸出履歴シートの関連レコードも更新
        if (originalId !== newId) {
            const rentalValues = rentalSheet.getDataRange().getValues();
            const idColIndex = HISTORY_HEADER.indexOf('管理番号');
            
            for (let i = 1; i < rentalValues.length; i++) { // ヘッダー行を除く
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
 * @param {string} itemId 管理番号
 */
function deleteItem(itemId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        // 貸出中の物品は削除不可
        const status = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
        if(status === '貸出中') throw new Error('貸出中の物品は削除できません。');

        // 予約が入っている物品は削除不可
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

// ----------------------------------------------------------------
// 貸出・返却・予約 処理
// ----------------------------------------------------------------

/**
 * 物品を貸し出す
 * @param {object} data - { itemId, borrower, siteName, startDate, returnDate, email }
 */
function lendItem(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, startDate, returnDate, email } = data;
        if (!borrower) throw new Error('貸与者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        // カテゴリが「通信機器」の場合のみ現場名を必須にする
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();
        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        // 校正日・入替日チェック
        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            const today = normalizeDate(new Date());
            if (normalizeDate(maintenanceDate) < today) {
                throw new Error(`この物品は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を過ぎているため、貸し出せません。`);
            }
        }
        
        // ステータスチェック
        const currentStatus = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
        if (['貸出中'].includes(currentStatus)) {
            throw new Error('この物品は現在貸出中のため、新たに貸し出すことはできません。');
        }
        
        // 期間重複チェック
        const checkResult = checkAvailability(itemId, new Date(startDate), new Date(returnDate));
        if (!checkResult.isAvailable) {
            throw new Error(`貸出期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        // 履歴シートに「貸出」レコードを追加
        addHistoryRecord(itemId, null, '貸出', data);
        // 物品マスタのステータスを更新
        updateItemStatus(itemId);

        // メール通知処理
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
test`;

            sendNotificationEmail(email, subject, body);
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
 * @param {string} historyId 履歴ID
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

        // 履歴シートのステータスを「返却済み」にし、実返却日を記録
        historyValues[HISTORY_HEADER.indexOf('処理')] = '返却済み';
        historyValues[HISTORY_HEADER.indexOf('実返却日')] = new Date();
        historyRange.setValues([historyValues]);

        // 物品マスタのステータスを更新
        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];
        updateItemStatus(itemId);

        // メール通知処理
        const email = historyValues[HISTORY_HEADER.indexOf('通知先アドレス')];
        if (email || (typeof mailSheet !== 'undefined' && mailSheet)) {
            const itemName = getItemNameById(itemId);
            const borrower = historyValues[HISTORY_HEADER.indexOf('貸与者')];
            const siteName = historyValues[HISTORY_HEADER.indexOf('現場名')];
            const startDate = historyValues[HISTORY_HEADER.indexOf('貸与日')];
            const returnDate = historyValues[HISTORY_HEADER.indexOf('実返却日')]; // 実返却日

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
kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
 * @param {object} data - { itemId, borrower, siteName, newStartDate, newReturnDate, email }
 */
function updateRentalPeriod(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, newStartDate, newReturnDate, email } = data;
        if (!borrower) throw new Error('貸与者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');

        // カテゴリが「通信機器」の場合のみ現場名を必須にする
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();
        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        // 校正日チェック
        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            if (normalizeDate(newReturnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`返却日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
        }

        // 更新対象の「貸出」履歴を探す
        const historyValues = rentalSheet.getDataRange().getValues();
        if (historyValues.length < 2) throw new Error('更新対象の貸出履歴が見つかりません。');
        
        // 最新（一番下）の「貸出」レコードを探す
        const historyRowIndex = historyValues.length - 1 - [...historyValues].reverse().findIndex(row => row[1] == itemId && row[3] === '貸出');
        if (historyRowIndex < 1) throw new Error('更新対象の貸出履歴が見つかりません。');
        
        // 重複チェック (自分自身の履歴IDを除外)
        const historyIdToIgnore = historyValues[historyRowIndex][0];
        const checkResult = checkAvailability(itemId, new Date(newStartDate), new Date(newReturnDate), true, historyIdToIgnore);
        if (!checkResult.isAvailable) {
            throw new Error(`変更後の期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        // 物品マスタの貸出情報を更新
        itemSheet.getRange(itemRow, HEADER.indexOf('貸与者') + 1).setValue(borrower);
        itemSheet.getRange(itemRow, HEADER.indexOf('現場名') + 1).setValue(siteName);
        itemSheet.getRange(itemRow, HEADER.indexOf('貸与日') + 1).setValue(Utilities.formatDate(normalizeDate(newStartDate), JST, 'yyyy/MM/dd'));
        itemSheet.getRange(itemRow, HEADER.indexOf('返却予定日') + 1).setValue(Utilities.formatDate(normalizeDate(newReturnDate), JST, 'yyyy/MM/dd'));
        
        // 貸出履歴シートの該当レコードを更新
        const historyRow = historyRowIndex + 1;
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('貸与者') + 1).setValue(borrower);
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('現場名') + 1).setValue(siteName);
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('貸与日') + 1).setValue(normalizeDate(newStartDate));
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('返却予定日') + 1).setValue(normalizeDate(newReturnDate));
        rentalSheet.getRange(historyRow, HISTORY_HEADER.indexOf('通知先アドレス') + 1).setValue(email || '');
        
        // メール通知処理
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
Kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
 * @param {object} data - { itemId, borrower, siteName, startDate, returnDate, email }
 */
function addReservation(data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { itemId, borrower, siteName, startDate, returnDate, email } = data;
        if (!borrower) throw new Error('予約者が入力されていません。');

        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。');
        
        // カテゴリが「通信機器」の場合のみ現場名を必須にする
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();
        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        // 校正日チェック
        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            const today = normalizeDate(new Date());
            // 予約終了日が校正日を過ぎていたらNG
            if (normalizeDate(returnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`予約終了日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
            if (normalizeDate(maintenanceDate) < today) {
                throw new Error(`この物品は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を過ぎているため、予約できません。`);
            }
        }

        // 期間重複チェック
        const checkResult = checkAvailability(itemId, new Date(startDate), new Date(returnDate));
        if (!checkResult.isAvailable) {
            throw new Error(`予約期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        // 履歴シートに「予約」レコードを追加
        addHistoryRecord(itemId, null, '予約', data);
        // 物品マスタのステータスを更新
        updateItemStatus(itemId);

        // メール通知処理
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
Kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
 * @param {object} data - { historyId, borrower, siteName, newStartDate, newReturnDate, email }
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

        // 物品マスタでカテゴリと校正日をチェック
        const itemRow = findRowById(itemSheet, itemId);
        if (itemRow === -1) throw new Error('対象の物品が見つかりません。'); // 物品が削除されたケース

        // カテゴリが「通信機器」の場合のみ現場名を必須にする
        const itemCategory = itemSheet.getRange(itemRow, HEADER.indexOf('カテゴリ') + 1).getValue();
        if (itemCategory === '通信機器' && !siteName) {
            throw new Error('「通信機器」カテゴリの物品には、現場名の入力が必須です。');
        }

        // 校正日チェック
        const itemDataValues = itemSheet.getRange(itemRow, 1, 1, HEADER.length).getValues()[0];
        const maintenanceDate = itemDataValues[HEADER.indexOf('校正日・入替日')];
        if (maintenanceDate instanceof Date) {
            if (normalizeDate(newReturnDate) > normalizeDate(maintenanceDate)) {
                throw new Error(`予約終了日は校正・入替期限(${Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')})を超えることはできません。`);
            }
        }

        // 期間重複チェック (自分自身の履歴IDを除外)
        const checkResult = checkAvailability(itemId, new Date(newStartDate), new Date(newReturnDate), false, historyId);
        if (!checkResult.isAvailable) {
            throw new Error(`変更後の期間が他の予定と重複しています。\n重複期間: ${checkResult.conflictReason}`);
        }

        // 履歴シートの該当レコードを更新
        historyValues[HISTORY_HEADER.indexOf('貸与者')] = borrower;
        historyValues[HISTORY_HEADER.indexOf('現場名')] = siteName;
        historyValues[HISTORY_HEADER.indexOf('貸与日')] = normalizeDate(newStartDate);
        historyValues[HISTORY_HEADER.indexOf('返却予定日')] = normalizeDate(newReturnDate);
        historyValues[HISTORY_HEADER.indexOf('通知先アドレス')] = email || '';
        
        historyRange.setValues([historyValues]);
        
        // メール通知処理
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
Kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
 * @param {string} historyId 履歴ID
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

        // 履歴シートのステータスを「キャンセル済」に変更
        historyValues[HISTORY_HEADER.indexOf('処理')] = 'キャンセル済';
        historyRange.setValues([historyValues]);
        
        // 物品マスタのステータスを更新
        updateItemStatus(itemId);
        
        // メール通知処理
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
Kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
 * 予約を「貸出」ステータスに変更する（カレンダー詳細モーダルからの手動操作）
 * @param {string} historyId 履歴ID
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

        // 予約開始日より前は貸出不可
        if (startDate > today) {
            throw new Error(`この予約は${Utilities.formatDate(startDate, JST, 'M月d日')}開始のため、まだ貸し出せません。`);
        }

        // 履歴シートのステータスを「貸出」に変更
        historyValues[HISTORY_HEADER.indexOf('処理')] = '貸出';
        historyRange.setValues([historyValues]);
        
        const itemId = historyValues[HISTORY_HEADER.indexOf('管理番号')];
        const email = historyValues[HISTORY_HEADER.indexOf('通知先アドレス')];
        updateItemStatus(itemId); // 物品マスタのステータス更新
        
        // メール通知処理
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
Kounai_rental`;

            sendNotificationEmail(email, subject, body);
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
    // 1. 貸出・予約・返却済みイベントの取得
    const rentalLastRow = rentalSheet.getLastRow();
    const rentalValues = rentalLastRow < 2 ? [] : rentalSheet.getRange(2, 1, rentalLastRow - 1, HISTORY_HEADER.length).getValues();

    const rentalEvents = rentalValues
        .filter(row => ['貸出', '予約', '返却済み'].includes(row[3])) // 該当ステータスのみ
        .map(row => {
            const history = formatRowAsObject(row, HISTORY_HEADER);
            const startDate = normalizeDate(history.貸与日);
            
            // 返却済みの場合、実返却日を終了日として採用
            let effectiveEndDate = normalizeDate(history.返却予定日);
            if (history.処理 === '返却済み' && history.実返却日) {
                effectiveEndDate = normalizeDate(history.実返却日);
            }

            // FullCalendarは終了日の翌日を指定する必要がある
            const endForCalendar = new Date(effectiveEndDate.getTime());
            endForCalendar.setDate(endForCalendar.getDate() + 1);

            let eventData = {
                id: history.履歴ID,
                resourceId: history.管理番号,
                start: Utilities.formatDate(startDate, JST, 'yyyy-MM-dd'),
                end: Utilities.formatDate(endForCalendar, JST, 'yyyy-MM-dd'),
                extendedProps: { // モーダル表示用の追加情報
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

            // ステータス別に色とタイトルを設定
            switch(history.処理) {
                case '貸出':
                    eventData.title = `${history.現場名 || history.貸与者} (貸出)`;
                    eventData.backgroundColor = '#d9534f'; // 赤系
                    eventData.borderColor = '#d43f3a';
                    eventData.extendedProps.type = 'rental';
                    break;
                case '予約':
                    eventData.title = `${history.現場名 || history.貸与者} (予約)`;
                    eventData.backgroundColor = '#f0ad4e'; // オレンジ系
                    eventData.borderColor = '#eea236';
                    eventData.extendedProps.type = 'reservation';
                    break;
                case '返却済み':
                    eventData.title = `【返却済み】${history.現場名 || history.貸与者}`;
                    eventData.backgroundColor = '#adb5bd'; // グレー系
                    eventData.borderColor = '#6c757d';
                    eventData.extendedProps.type = 'returned';
                    break;
            }
            return eventData;
        });

    // 2. 校正日・入替日イベントの取得
    const itemLastRow = itemSheet.getLastRow();
    const itemValues = itemLastRow < 2 ? [] : itemSheet.getRange(2, 1, itemLastRow - 1, HEADER.length).getValues();
    
    const maintenanceEvents = itemValues
        .map(row => formatRowAsObject(row, HEADER))
        .filter(item => item['校正日・入替日']) // 校正日があるもののみ
        .map(item => {
            const maintenanceDate = normalizeDate(item['校正日・入替日']);
            return {
                id: `maint_${item.管理番号}`,
                resourceId: item.管理番号,
                title: `校正/入替: ${item.物品名}`,
                start: Utilities.formatDate(maintenanceDate, JST, 'yyyy-MM-dd'),
                allDay: true,
                backgroundColor: '#6c757d', // グレー系
                borderColor: '#6c757d',
                extendedProps: {
                    type: 'maintenance',
                    itemId: item.管理番号,
                    itemName: item.物品名,
                    maintenanceDate: Utilities.formatDate(maintenanceDate, JST, 'yyyy/MM/dd')
                }
            };
        });
    
    return [...rentalEvents, ...maintenanceEvents]; // すべてのイベントを結合
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
    if (!lock.tryLock(300000)) { // 5分待機
        Logger.log("他のプロセスが実行中のため、スキップします。");
        return;
    }

    try {
        const today = normalizeDate(new Date());

        if (rentalSheet.getLastRow() < 2) {
            Logger.log("処理対象の予約はありません。");
            return;
        }
        const events = rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, HISTORY_HEADER.length).getValues();

        events.forEach((event, index) => {
            const status = event[3];
            const startDate = normalizeDate(event[6]);
            
            // ステータスが「予約」で、開始日が「今日」
            if (status === '予約' && startDate.getTime() === today.getTime()) {
                const itemId = event[1];
                const itemRow = findRowById(itemSheet, itemId);
                
                if (itemRow !== -1) {
                    // 物品マスタのステータスが「予約あり」（＝現在貸出中でない）ことを確認
                    const itemStatus = itemSheet.getRange(itemRow, HEADER.indexOf('ステータス') + 1).getValue();
                    if (itemStatus === '予約あり') {
                        Logger.log(`予約ID[${event[0]}]を自動で貸し出します。`);
                        rentalSheet.getRange(index + 2, HISTORY_HEADER.indexOf('処理') + 1).setValue('貸出');
                        SpreadsheetApp.flush(); // 即時反映
                        updateItemStatus(itemId); // 物品マスタも更新
                    } else {
                        // 貸出中のまま返却されていない場合
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
 * @param {string} itemId 管理番号
 * @param {string} itemName 物品名 (null許容)
 * @param {string} action 処理 ('貸出', '予約' など)
 * @param {object} data { borrower, siteName, startDate, returnDate, email }
 */
function addHistoryRecord(itemId, itemName, action, data) {
    const id = getNextId(rentalSheet); // 新しい履歴IDを採番
    const name = itemName || (itemId ? getItemNameById(itemId) : '不明な物品'); 
    const newRow = [
        id, itemId, name, action, data.borrower,
        data.siteName || '',
        normalizeDate(data.startDate), normalizeDate(data.returnDate),
        '', // 実返却日 (空)
        data.email || '' // 通知先アドレス
    ];
    rentalSheet.appendRow(newRow);
}

/**
 * 物品マスタのステータスと貸出情報を最新の状態に更新する
 * @param {string} itemId 管理番号
 */
function updateItemStatus(itemId) {
    const itemRow = findRowById(itemSheet, itemId);
    if (itemRow === -1) return;

    const itemRange = itemSheet.getRange(itemRow, 1, 1, HEADER.length);
    const itemValues = itemRange.getValues()[0];
    const allFutureEvents = getFutureEventsForItem(itemId);
    const today = normalizeDate(new Date());
    
    // 現在アクティブな「貸出」を探す
    const activeRental = allFutureEvents.find(e => e[3] === '貸出');

    if (activeRental) {
        // 貸出中の場合
        itemValues[HEADER.indexOf('ステータス')] = '貸出中';
        itemValues[HEADER.indexOf('貸与者')] = activeRental[4];
        itemValues[HEADER.indexOf('現場名')] = activeRental[5];
        
        // 日付をJSTのyyyy/MM/dd形式の文字列として保存
        const lendDate = activeRental[6];
        const formattedLendDate = Utilities.formatDate(lendDate, JST, 'yyyy/MM/dd');
        
        itemValues[HEADER.indexOf('貸与日')] = formattedLendDate;
        itemValues[HEADER.indexOf('返却予定日')] = Utilities.formatDate(activeRental[7], JST, 'yyyy/MM/dd');

    } else {
        // 貸出中でない場合
        // 今日以降の「予約」を探す
        const upcomingReservation = allFutureEvents.find(e => 
            e[3] === '予約' && normalizeDate(e[6]) >= today
        );
        
        itemValues[HEADER.indexOf('ステータス')] = upcomingReservation ? '予約あり' : '在庫';
        // 貸出情報をクリア
        itemValues[HEADER.indexOf('貸与者')] = '';
        itemValues[HEADER.indexOf('現場名')] = '';
        itemValues[HEADER.indexOf('貸与日')] = '';
        itemValues[HEADER.indexOf('返却予定日')] = '';
    }
    itemRange.setValues([itemValues]); // 物品マスタの行を更新
}

/**
 * 特定の物品の、今日以降に終了する「貸出」または「予約」の履歴を取得する
 * @param {string} itemId 管理番号
 * @return {Array[]} 履歴データの配列
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
            // 返却予定日(eventEndDate)が今日以降
            const eventEndDate = normalizeDate(row[7]);
            return isTargetItem && isFutureEvent && eventEndDate >= today;
        })
        .sort((a,b) => normalizeDate(a[6]) - normalizeDate(b[6])); // 開始日順にソート
}

/**
 * 指定された期間が、他の既存の予定と重複していないかチェックする
 * @param {string} itemId 管理番号
 * @param {Date} newStart 新しい開始日
 * @param {Date} newEnd 新しい終了日
 * @param {boolean} isEditingLend (未使用)
 * @param {string} historyIdToIgnore 無視する履歴ID (編集中の予定自体)
 * @return {object} { isAvailable: boolean, conflictReason: string }
 */
function checkAvailability(itemId, newStart, newEnd, isEditingLend = false, historyIdToIgnore = null) {
    const normalizedStart = normalizeDate(newStart);
    const normalizedEnd = normalizeDate(newEnd);

    const futureEvents = getFutureEventsForItem(itemId);

    for (const event of futureEvents) {
        // 自分自身（編集中の予定）はチェック対象から除外
        if (String(event[0]) === String(historyIdToIgnore)) continue;

        const eventStart = normalizeDate(event[6]);
        const eventEnd = normalizeDate(event[7]);
        
        // 期間重複ロジック (A_start <= B_end AND A_end >= B_start)
        if (normalizedStart <= eventEnd && normalizedEnd >= eventStart) {
            const reason = `${event[3]}あり (${Utilities.formatDate(eventStart, JST, 'yyyy/MM/dd')} ~ ${Utilities.formatDate(eventEnd, JST, 'yyyy/MM/dd')})`;
            return { isAvailable: false, conflictReason: reason };
        }
    }
    return { isAvailable: true }; // 重複なし
}

/**
 * 指定されたシートのA列からIDを検索し、その行番号（1始まり）を返す
 * @param {SpreadsheetApp.Sheet} sheet 検索対象シート
 * @param {string} id 検索するID
 * @return {number} 行番号 (見つからない場合は -1)
 */
function findRowById(sheet, id) {
    if (!id || sheet.getLastRow() < 2) return -1;
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(cellId => String(cellId) == String(id)); // 文字列として比較
    return (rowIndex !== -1) ? rowIndex + 2 : -1; // 0-based indexを 1-based row number (ヘッダー+1) に変換
}

/**
 * スプレッドシートの行データ（配列）を、ヘッダーに基づいたオブジェクトに変換する
 * 日付は 'yyyy/MM/dd' 形式の文字列に変換する
 * @param {Array} rowArray 行データの配列
 * @param {Array} headerArray ヘッダー定義の配列
 * @return {object} 変換後のオブジェクト
 */
function formatRowAsObject(rowArray, headerArray) {
    let item = {};
    const DATE_KEYS = ['貸出日', '返却予定日', '校正日・入替日', '貸与日', '実返却日']; 
    headerArray.forEach((key, index) => {
        const value = rowArray[index];
        
        if (DATE_KEYS.includes(key)) {
            // 日付キーの場合、正規化してフォーマット
            const normalized = normalizeDate(value);
            if (normalized) {
                item[key] = Utilities.formatDate(normalized, JST, 'yyyy/MM/dd');
            } else {
                item[key] = ''; // 不正な日付や空欄は空文字列
            }
        } else {
            item[key] = value;
        }
    });
    return item;
}

/**
 * 貸出履歴シートのA列の最大ID+1の値を返す
 * @param {SpreadsheetApp.Sheet} sheet 貸出履歴シート
 * @return {number} 次のID
 */
function getNextId(sheet) {
    const lastRow = rentalSheet.getLastRow();
    if (lastRow < 2) return 1; // データがない場合は1
    const ids = rentalSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const maxId = Math.max(0, ...ids.filter(id => !isNaN(id)).map(id => Number(id)));
    return isFinite(maxId) ? maxId + 1 : 1;
}

/**
 * タイムゾーンのズレを吸収し、JSTの0時0分0秒のDateオブジェクトを返す
 * @param {Date | string | number} date - 変換対象の日付
 * @return {Date | null} JSTの0時0分0秒に正規化されたDateオブジェクト
 */
function normalizeDate(date) {
    if (!date) return null;

    try {
        let d;
        if (date instanceof Date) {
            // JSTの「yyyy/MM/dd」文字列を一度取得する
            const dateString = Utilities.formatDate(date, JST, 'yyyy/MM/dd');
            d = new Date(dateString); // 'yyyy/MM/dd' 形式からDateを生成するとJST 0時になる
        } 
        else if (typeof date === 'string') {
            // 'yyyy-MM-dd' や 'yyyy/MM/dd' 形式を想定
            const dateString = date.split('T')[0].replace(/-/g, '/');
            d = new Date(dateString); // 'yyyy/MM/dd' 形式からDateを生成するとJST 0時になる
        } 
        else {
            // 数値（タイムスタンプ）など
            const tempDate = new Date(date);
            const dateString = Utilities.formatDate(tempDate, JST, 'yyyy/MM/dd');
            d = new Date(dateString);
        }
        
        if (isNaN(d.getTime())) { // 不正な日付チェック
            return null;
        }

        return d; // JSTの0時0分0秒のDateオブジェクト
    } catch(e) {
        Logger.log("normalizeDate Error: " + e.message + " (Input: " + date + ")");
        return null;
    }
}

// ----------------------------------------------------------------
// メール送信用ヘルパー関数
// ----------------------------------------------------------------

/**
 * メールマスタシートから指定された種類のメールアドレスを取得する
 * @param {string} mailType - メール種類 (例: '編集')
 * @return {string[]} メールアドレスの配列
 */
function getMailMasterAddresses(mailType) {
    if (typeof mailSheet === 'undefined' || !mailSheet) {
        Logger.log("「メールマスタ」シートが見つかりません。");
        return [];
    }
    try {
        const lastRow = mailSheet.getLastRow();
        if (lastRow < 2) return [];

        // A列: 送り先, B列: メール種類, C列: アドレス と想定
        const values = mailSheet.getRange(2, 1, lastRow - 1, 3).getValues();
        
        const addresses = values
            .filter(row => row[1] === mailType && row[2]) // B列がmailTypeと一致し、C列にアドレスがある
            .map(row => row[2].trim()) // C列のアドレス
            .filter(address => address.length > 0 && address.includes('@')); // 有効なアドレスか簡易チェック

        return [...new Set(addresses)]; // 重複を除外
    } catch (e) {
        Logger.log(`getMailMasterAddresses Error: ${e.message} ${e.stack}`);
        return [];
    }
}

/**
 * 操作完了通知メールを送信する
 * @param {string} formEmail - フォームから入力されたメールアドレス (カンマ区切り)
 * @param {string} subject - メールの件名
 * @param {string} body - メールの本文
 */
function sendNotificationEmail(formEmail, subject, body) {
    let recipients = [];

    // 1. フォームからのアドレスを追加
    if (formEmail) {
        const emailsFromForm = formEmail.split(',')
            .map(email => email.trim())
            .filter(email => email.length > 0 && email.includes('@'));
        recipients = recipients.concat(emailsFromForm);
    }

    // 2. メールマスタ（'編集'）からのアドレスを追加
    const masterAddresses = getMailMasterAddresses('編集');
    recipients = recipients.concat(masterAddresses);

    // 3. 宛先の重複を除外し、空でなければ送信
    const uniqueRecipients = [...new Set(recipients)].filter(email => email);

    if (uniqueRecipients.length > 0) {
        try {
            const recipientString = uniqueRecipients.join(',');
            MailApp.sendEmail({
                to: recipientString,
                subject: subject,
                body: body,
                // noReply: true // 返信不可にする場合はコメントアウトを解除
            });
            Logger.log(`メール送信成功: ${recipientString}`);
        } catch (e) {
            Logger.log(`メール送信エラー: ${e.message} ${e.stack}`);
            // メール送信エラーは操作の成功/失敗に影響させない
        }
    } else {
        Logger.log("通知先アドレスが指定されていないため、メールは送信されませんでした。");
    }
}

/**
 * 物品IDから物品名を取得する（メール本文用）
 * @param {string} itemId
 * @return {string} 物品名 (見つからない場合は '不明な物品(ID: ...)')
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
