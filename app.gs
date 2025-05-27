// 設置網頁應用
function doGet(e) {
  // 一般的網頁應用入口點
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('案件進度追蹤系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 全局變數

// API 函數，用於前端調用
function getAllCases(sheetsId, keyword) {
  try {
    Logger.log('====== 開始執行 getAllCases 函數 ======');
    Logger.log('參數 - sheets ID: ' + sheetsId + ', 關鍵字: ' + keyword);
    
    if (!sheetsId) {
      Logger.log('錯誤：未提供 Google Sheets ID');
      throw new Error('未提供 Google Sheets ID');
    }
    
    let ss;
    try {
      Logger.log('嘗試開啟試算表，ID: ' + sheetsId);
      ss = SpreadsheetApp.openById(sheetsId);
      Logger.log('成功開啟試算表: ' + ss.getName());
    } catch (sheetError) {
      Logger.log('打開試算表時發生錯誤，完整錯誤信息: ' + JSON.stringify(sheetError));
      console.error('打開試算表時發生錯誤:', sheetError);
      console.error('錯誤詳情:', sheetError.stack);
      throw new Error('無法打開指定的試算表: ' + sheetError.toString() + '。可能原因：權限不足或ID無效');
    }
    
    const casesSheet = ss.getSheetByName('案件');
    const progressSheet = ss.getSheetByName('進度');
    const relationsSheet = ss.getSheetByName('相關人關聯');
    const peopleSheet = ss.getSheetByName('相關人');
    const categoriesSheet = ss.getSheetByName('類別');
    const categoryRelationsSheet = ss.getSheetByName('類別關聯');
    
    if (!casesSheet || !progressSheet || !relationsSheet || !categoriesSheet || !categoryRelationsSheet) {
      // 嘗試初始化工作表
      Logger.log('找不到必要的工作表，嘗試初始化...');
      console.log('找不到必要的工作表，嘗試初始化...');
      initializeSheets(sheetsId);
      return { 
        success: true, 
        cases: [], 
        completedCases: [],
        people: [],
        categories: [],
        message: '已初始化工作表，請重新載入資料'
      };
    }
    
    // 讀取案件資料
    let casesData, progressData, relationsData, peopleData, categoriesData, categoryRelationsData;
    try {
      console.log('開始讀取案件資料...');
      Logger.log('開始讀取案件資料，工作表名稱: ' + casesSheet.getName());
      casesData = getSheetData(casesSheet);
      console.log(`成功讀取 ${casesData.length} 筆案件資料`);
      Logger.log(`成功讀取 ${casesData.length} 筆案件資料`);
      
      // 記錄原始讀取的案件資料 (前5筆)
      if (casesData && casesData.length > 0) {
        Logger.log(`讀取的原始案件資料 (前${Math.min(5, casesData.length)}筆): ${JSON.stringify(casesData.slice(0, 5))}`);
      }
      
      console.log('開始讀取進度資料...');
      Logger.log('開始讀取進度資料，工作表名稱: ' + progressSheet.getName());
      progressData = getSheetData(progressSheet);
      console.log(`成功讀取 ${progressData.length} 筆進度資料`);
      Logger.log(`成功讀取 ${progressData.length} 筆進度資料`);
      
      // 記錄原始讀取的進度資料 (前5筆)
      if (progressData && progressData.length > 0) {
        Logger.log(`讀取的原始進度資料 (前${Math.min(5, progressData.length)}筆): ${JSON.stringify(progressData.slice(0, 5))}`);
      }
      
      // 讀取相關人關聯資料
      console.log('開始讀取相關人關聯資料...');
      Logger.log('開始讀取相關人關聯資料...');
      relationsData = getSheetData(relationsSheet);
      console.log(`成功讀取 ${relationsData.length} 筆相關人關聯資料`);
      Logger.log(`成功讀取 ${relationsData.length} 筆相關人關聯資料`);
      
      // 記錄原始相關人關聯資料 (前10筆)
      if (relationsData && relationsData.length > 0) {
        Logger.log(`讀取的原始相關人關聯資料 (前${Math.min(10, relationsData.length)}筆): ${JSON.stringify(relationsData.slice(0, 10))}`);
      }
      
      // 讀取相關人資料
      console.log('開始讀取相關人資料...');
      Logger.log('開始讀取相關人資料...');
      if (peopleSheet) {
        peopleData = getSheetData(peopleSheet);
        console.log(`成功讀取 ${peopleData.length} 筆相關人資料`);
        Logger.log(`成功讀取 ${peopleData.length} 筆相關人資料`);
        
        // 記錄原始相關人資料
        Logger.log(`讀取的原始相關人資料 (全部): ${JSON.stringify(peopleData)}`);
      } else {
        console.log('相關人工作表不存在，將創建它');
        Logger.log('相關人工作表不存在，將創建它');
        initializeSheets(sheetsId);
        peopleData = [];
      }
      
      // 讀取類別資料
      console.log('開始讀取類別資料...');
      Logger.log('開始讀取類別資料...');
      if (categoriesSheet) {
        categoriesData = getSheetData(categoriesSheet);
        console.log(`成功讀取 ${categoriesData.length} 筆類別資料`);
        Logger.log(`成功讀取 ${categoriesData.length} 筆類別資料`);
        
        // 記錄原始類別資料
        Logger.log(`讀取的原始類別資料 (全部): ${JSON.stringify(categoriesData)}`);

        // 記錄每個類別的原始順序值
        categoriesData.forEach(cat => {
          Logger.log(`類別 ${cat.id} (${cat.name}) 的原始 order 值: ${cat.order}, 類型: ${typeof cat.order}`);
        });
      } else {
        console.log('類別工作表不存在，將創建它');
        Logger.log('類別工作表不存在，將創建它');
        initializeSheets(sheetsId);
        categoriesData = [];
      }
      
      // 讀取類別關聯資料
      console.log('開始讀取類別關聯資料...');
      Logger.log('開始讀取類別關聯資料...');
      if (categoryRelationsSheet) {
        categoryRelationsData = getSheetData(categoryRelationsSheet);
        console.log(`成功讀取 ${categoryRelationsData.length} 筆類別關聯資料`);
        Logger.log(`成功讀取 ${categoryRelationsData.length} 筆類別關聯資料`);
        
        // 記錄原始類別關聯資料
        Logger.log(`讀取的原始類別關聯資料 (全部): ${JSON.stringify(categoryRelationsData)}`);
      } else {
        console.log('類別關聯工作表不存在，將創建它');
        Logger.log('類別關聯工作表不存在，將創建它');
        initializeSheets(sheetsId);
        categoryRelationsData = [];
      }

    } catch (dataError) {
      Logger.log('讀取工作表資料時發生錯誤: ' + JSON.stringify(dataError));
      Logger.log('錯誤堆疊: ' + dataError.stack);
      console.error('讀取工作表資料時發生錯誤:', dataError);
      console.error('錯誤堆疊:', dataError.stack);
      throw new Error('讀取工作表資料時發生錯誤: ' + dataError.toString() + '。詳細資訊: ' + (dataError.stack || '無堆疊信息'));
    }
    
    // 格式化案件資料
    const cases = [];
    const completedCases = [];
    
    // 格式化相關人資料
    const people = peopleData.map(person => {
      return {
        id: person.id,
        name: person.name,
        color: person.color || generateRandomColor(),
        createDate: person.createDate || new Date().toISOString(),
        isActive: !(person.isActive === 'false' || person.isActive === false)
      };
    });
    
    // 記錄格式化後的相關人資料
    Logger.log(`格式化後的相關人資料 (全部): ${JSON.stringify(people)}`);
    Logger.log(`活躍相關人數量: ${people.filter(p => p.isActive).length}, 非活躍相關人數量: ${people.filter(p => !p.isActive).length}`);
    
    // 格式化類別資料並正確排序
    let categories = categoriesData.map(category => {
      // 確保order是數字類型
      let orderValue = category.order;
      if (typeof orderValue === 'string') {
        orderValue = Number(orderValue);
        if (isNaN(orderValue)) orderValue = 9999;
        Logger.log(`[getAllCases] 類別 ${category.id} (${category.name}) 的 order 從字串 "${category.order}" 轉換為數字 ${orderValue}`);
      } else if (orderValue === undefined || orderValue === null) {
        orderValue = 9999;
        Logger.log(`[getAllCases] 類別 ${category.id} (${category.name}) 沒有 order 值，設為預設值 ${orderValue}`);
      } else {
        orderValue = Number(orderValue);
        Logger.log(`[getAllCases] 類別 ${category.id} (${category.name}) 的 order 值是 ${orderValue}, 轉換前類型: ${typeof category.order}`);
      }
      
      return {
        id: category.id,
        name: category.name,
        color: category.color || generateRandomColor(),
        createDate: category.createDate || new Date().toISOString(),
        isActive: !(category.isActive === 'false' || category.isActive === false),
        order: orderValue // 使用轉換後的數字
      };
    });
    
    // 按 order 欄位排序前的數據
    Logger.log(`[getAllCases] 排序前的類別數據: ${JSON.stringify(categories.map(c => ({id: c.id, name: c.name, order: c.order, orderType: typeof c.order})))}`);
    
    // 強制確保所有 order 值都是數字
    for (let i = 0; i < categories.length; i++) {
      if (typeof categories[i].order !== 'number') {
        const oldValue = categories[i].order;
        categories[i].order = Number(categories[i].order) || 9999;
        Logger.log(`[getAllCases] 強制修正：類別 ${categories[i].id} (${categories[i].name}) 的 order 從 ${oldValue} (${typeof oldValue}) 修正為 ${categories[i].order}`);
      }
    }
    
    // 簡化排序邏輯，直接使用數字比較
    categories.sort((a, b) => {
      // 確保 a.order 和 b.order 都是數字
      const orderA = Number(a.order);
      const orderB = Number(b.order);
      
      Logger.log(`[getAllCases] 排序比較: ${a.name}(order=${orderA}, ${typeof orderA}) vs ${b.name}(order=${orderB}, ${typeof orderB}), 結果: ${orderA - orderB}`);
      
      return orderA - orderB;
    });
    
    // 排序後的數據
    Logger.log(`[getAllCases] 排序後的類別數據: ${JSON.stringify(categories.map(c => ({id: c.id, name: c.name, order: c.order})))}`);
    
    // 再次確認所有類別的順序值
    Logger.log('[getAllCases] 最終排序後的每個類別順序:');
    for (let i = 0; i < categories.length; i++) {
      Logger.log(`[getAllCases][${i}] 類別 ${categories[i].id} (${categories[i].name}): order = ${categories[i].order}, 類型: ${typeof categories[i].order}`);
    }
    
    // 記錄格式化後的類別資料
    Logger.log(`格式化後的類別資料 (全部): ${JSON.stringify(categories)}`);
    Logger.log(`活躍類別數量: ${categories.filter(c => c.isActive).length}, 非活躍類別數量: ${categories.filter(c => !c.isActive).length}`);
    
    // --- 關鍵字篩選邏輯 --- 
    const lowerKeyword = keyword ? keyword.toLowerCase() : null;
    let filteredCasesData = casesData;

    if (lowerKeyword) {
      Logger.log('開始執行關鍵字篩選...');
      filteredCasesData = casesData.filter(caseRow => {
        // 檢查標題和內容
        if (caseRow.title && caseRow.title.toLowerCase().includes(lowerKeyword)) return true;
        if (caseRow.content && caseRow.content.toLowerCase().includes(lowerKeyword)) return true;
        
        // 檢查進度內容
        const relatedProgress = progressData.filter(p => p.caseId === caseRow.id);
        for (const progressRow of relatedProgress) {
          if (progressRow.content && progressRow.content.toLowerCase().includes(lowerKeyword)) return true;
        }
        
        return false; // 如果以上都不符合，則過濾掉
      });
      Logger.log(`關鍵字篩選後，剩下 ${filteredCasesData.length} 筆案件`);
    }
    // --- 篩選邏輯結束 ---

    // 使用篩選後的 filteredCasesData 進行格式化
    for (let i = 0; i < filteredCasesData.length; i++) {
      try {
        const caseRow = filteredCasesData[i];
        
        // 基本數據驗證
        if (!caseRow.id || !caseRow.title) {
          console.warn(`跳過無效的案件數據(缺少必要欄位) - 行 ${i+1}`);
          continue;
        }
        
        // 記錄當前處理的案件ID
        Logger.log(`====== 開始處理案件 ID: ${caseRow.id}, 標題: ${caseRow.title} ======`);
        
        // 從關聯表中獲取該案件的相關人ID
        const caseRelations = relationsData.filter(relation => 
          relation.objectId === caseRow.id && relation.objectType === '案件');
        const relatedPeopleFromRelations = caseRelations.map(relation => relation.personId);
        
        // 記錄從關聯表獲取的相關人ID
        Logger.log(`[案件 ID: ${caseRow.id}] 從關聯表獲取的相關人ID: ${JSON.stringify(relatedPeopleFromRelations)}`);
        
        // 檢查案件資料中是否有直接存儲的相關人資訊
        let relatedPeople = relatedPeopleFromRelations;
        if (caseRow.relatedPeople && typeof caseRow.relatedPeople === 'string' && caseRow.relatedPeople.trim() !== '') {
          try {
            Logger.log(`[案件 ID: ${caseRow.id}] G欄位的相關人原始字串: "${caseRow.relatedPeople}"`);
            const parsedRelatedPeople = JSON.parse(caseRow.relatedPeople);
            Logger.log(`[案件 ID: ${caseRow.id}] 從G欄位解析出的相關人資訊: ${JSON.stringify(parsedRelatedPeople)}, 類型: ${Array.isArray(parsedRelatedPeople) ? '陣列' : typeof parsedRelatedPeople}`);
            
            if (Array.isArray(parsedRelatedPeople) && parsedRelatedPeople.length > 0) {
              Logger.log(`[案件 ID: ${caseRow.id}] 使用G欄位中的相關人資訊，替換關聯表中的資訊`);
              relatedPeople = parsedRelatedPeople; // 優先使用 G 欄位的相關人資訊
            } else {
              Logger.log(`[案件 ID: ${caseRow.id}] G欄位中的相關人資訊不是有效的陣列，將繼續使用關聯表中的資訊`);
            }
          } catch (jsonError) {
            Logger.log(`[案件 ID: ${caseRow.id}] 解析G欄位相關人資訊時發生錯誤: ${jsonError.toString()}, 將使用關聯表中的資訊`);
            // 發生錯誤時使用關聯表中的相關人資訊
          }
        } else {
          Logger.log(`[案件 ID: ${caseRow.id}] 案件沒有G欄位相關人資訊或格式不正確，將使用關聯表中的資訊`);
        }
        
        // 最終確認使用的相關人列表
        Logger.log(`[案件 ID: ${caseRow.id}] 最終確定使用的相關人ID: ${JSON.stringify(relatedPeople)}`);
        
        // 檢查每個相關人ID是否存在於人員列表中
        if (relatedPeople && relatedPeople.length > 0) {
          relatedPeople.forEach(personId => {
            const personExists = peopleData.some(p => p.id === personId);
            if (personExists) {
              const person = people.find(p => p.id === personId);
              Logger.log(`[案件 ID: ${caseRow.id}] 相關人 ID: ${personId} 在人員列表中找到，名稱: ${person ? person.name : '未知'}, 狀態: ${person && person.isActive ? '活躍' : '非活躍'}`);
            } else {
              Logger.log(`[案件 ID: ${caseRow.id}] 警告：相關人 ID: ${personId} 在人員列表中不存在`);
            }
          });
        }
        
        // === 新增偵錯記錄：檢查原始 files 字串 ===
        let filesArray = [];
        if (caseRow.files && typeof caseRow.files === 'string') {
          const rawFileString = caseRow.files;
          Logger.log(`[案件 ID: ${caseRow.id}] 準備解析 files 字串: "${rawFileString}" (長度: ${rawFileString.length})`);
          // === 結束新增偵錯記錄 ===
          
          try {
            filesArray = JSON.parse(rawFileString);
            // === 新增偵錯記錄：解析成功 ===
            Logger.log(`[案件 ID: ${caseRow.id}] files 字串解析成功，結果: ${JSON.stringify(filesArray)}`);
            // === 結束新增偵錯記錄 ===
          } catch (jsonError) {
            // === 修改偵錯記錄：解析失敗 ===
            Logger.log(`[案件 ID: ${caseRow.id}] 解析 files 字串時發生錯誤: ${jsonError.toString()}`);
            console.error(`[案件 ID: ${caseRow.id}] 解析 files 字串時發生錯誤:`, jsonError);
            // === 結束修改偵錯記錄 ===
            filesArray = []; // 使用空陣列作為後備
          }
        } else if (caseRow.files) {
          // 如果 caseRow.files 不是字串，記錄其類型和值
          Logger.log(`[案件 ID: ${caseRow.id}] files 欄位不是字串，類型: ${typeof caseRow.files}, 值: ${JSON.stringify(caseRow.files)}`);
        }
        // === 結束新增偵錯記錄 ===
        
        const caseItem = {
          id: caseRow.id,
          title: caseRow.title,
          content: caseRow.content || '',
          date: caseRow.date || new Date().toISOString(),
          relatedPeople: relatedPeople,
          files: filesArray, // 確保使用解析後的 filesArray
          progress: [],
          completed: caseRow.completed === 'true',
          isFavorite: String(caseRow.isFavorite).toLowerCase() === 'true', // 處理大小寫問題
          isHidden: caseRow.isHidden,
          actionNeeded: caseRow.actionNeeded,
          isDeleted: caseRow.isDeleted // 確保將isDeleted屬性傳遞給前端
        };
        
        // 記錄案件的刪除狀態
        Logger.log(`[案件 ID: ${caseRow.id}] 刪除狀態: ${caseRow.isDeleted}, 類型: ${typeof caseRow.isDeleted}`);
        
        // 添加進度資料
        if (progressData && progressData.length > 0) {
          Logger.log(`[案件 ID: ${caseRow.id}] 開始處理進度資料...`);
          let progressCount = 0;
          
          for (let j = 0; j < progressData.length; j++) {
            try {
              const progressRow = progressData[j];
              
              if (progressRow.caseId === caseItem.id) {
                progressCount++;
                Logger.log(`[案件 ID: ${caseRow.id}] 處理進度 ID: ${progressRow.id}`);
                
                // 從關聯表中獲取該進度的相關人ID
                const progressRelations = relationsData.filter(relation => 
                  relation.objectId === progressRow.id && relation.objectType === '進度');
                const progressRelatedPeople = progressRelations.map(relation => relation.personId);
                
                Logger.log(`[案件 ID: ${caseRow.id}, 進度 ID: ${progressRow.id}] 相關人 ID: ${JSON.stringify(progressRelatedPeople)}`);
                
                // === 對進度檔案也做類似的檢查 (如果需要，此處暫時省略以保持簡潔) ===
                let progressFilesArray = [];
                if (progressRow.files) {
                  try {
                    progressFilesArray = JSON.parse(progressRow.files);
                    Logger.log(`[案件 ID: ${caseRow.id}, 進度 ID: ${progressRow.id}] 解析檔案成功`);
                  } catch (jsonError) {
                    console.error(`解析進度 ${progressRow.id} 檔案JSON時發生錯誤:`, jsonError);
                    Logger.log(`[案件 ID: ${caseRow.id}, 進度 ID: ${progressRow.id}] 解析檔案失敗: ${jsonError.toString()}`);
                    progressFilesArray = []; // 使用空陣列作為後備
                  }
                }
                // === 結束進度檔案檢查 ===
                
                caseItem.progress.push({
                  id: progressRow.id,
                  date: progressRow.date || new Date().toISOString(),
                  relatedPeople: progressRelatedPeople,
                  content: progressRow.content || '',
                  files: progressFilesArray,
                  isDeleted: progressRow.isDeleted // 確保包含isDeleted屬性
                });
              }
            } catch (progressError) {
              console.error(`處理案件 ${caseItem.id} 的進度資料時發生錯誤:`, progressError);
              Logger.log(`[案件 ID: ${caseRow.id}] 處理進度時發生錯誤: ${progressError.toString()}`);
              // 繼續處理下一個進度項目
            }
          }
          Logger.log(`[案件 ID: ${caseRow.id}] 共處理了 ${progressCount} 個進度項目`);
        }
        
        // 分類為進行中或已完成案件
        if (caseItem.completed) {
          completedCases.push(caseItem);
          Logger.log(`案件 ID: ${caseRow.id} 被分類為 [已完成]`);
        } else {
          cases.push(caseItem);
          Logger.log(`案件 ID: ${caseRow.id} 被分類為 [進行中]`);
        }
        
        Logger.log(`====== 完成處理案件 ID: ${caseRow.id} ======`);
      } catch (caseError) {
        console.error(`處理第 ${i+1} 筆案件資料時發生錯誤:`, caseError);
        Logger.log(`處理第 ${i+1} 筆案件資料時發生錯誤: ${caseError.toString()}`);
        // 繼續處理下一筆資料
      }
    }
    
    console.log(`成功處理 ${cases.length} 筆進行中案件和 ${completedCases.length} 筆已完成案件`);
    Logger.log(`====== 完成處理所有案件 ======`);
    Logger.log(`進行中案件: ${cases.length} 筆, 已完成案件: ${completedCases.length} 筆`);
    
    // 返回前記錄最終物件
    const finalResult = { 
      success: true, 
      cases: cases, 
      completedCases: completedCases,
      people: people, // 注意：將傳回所有相關人（包括非活躍）以確保前端顯示正確
      categories: categories // 注意：將傳回所有類別（包括非活躍）以確保前端顯示正確
    };
    Logger.log('最終準備返回前端的物件結構 (未序列化): ' + JSON.stringify({
      success: finalResult.success,
      casesCount: finalResult.cases.length,
      completedCasesCount: finalResult.completedCases.length,
      peopleCount: finalResult.people.length,
      categoriesCount: finalResult.categories.length
    }));
    
    // 將結果序列化為 JSON 字串返回
    return JSON.stringify(finalResult);
  } catch (error) {
    console.error('獲取案件時發生錯誤:', error);
    Logger.log('getAllCases 發生錯誤: ' + error.toString() + ' 堆疊: ' + error.stack);
    
    const errorResult = { 
      success: false, 
      error: error.toString() 
    };
    Logger.log('最終準備返回前端的錯誤物件 (序列化前): ' + JSON.stringify(errorResult));
    
    // 將錯誤結果也序列化返回
    return JSON.stringify(errorResult);
  }
}

// 新增案件
function addCase(sheetsId, driveFolderId, caseData) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!driveFolderId) {
      throw new Error('未提供 Google Drive 資料夾 ID');
    }
    
    Logger.log(`[addCase] 開始處理新增案件請求，標題: ${caseData.title}`);
    Logger.log(`[addCase] 收到的相關人資料: ${JSON.stringify(caseData.relatedPeople)}, 類型: ${typeof caseData.relatedPeople}`);
    Logger.log(`[addCase] 收到的類別資料: ${JSON.stringify(caseData.categories)}, 類型: ${typeof caseData.categories}`);
    
    // 確保 caseData.relatedPeople 是陣列
    if (caseData.relatedPeople && !Array.isArray(caseData.relatedPeople)) {
      Logger.log(`[addCase] 相關人數據不是陣列，嘗試轉換`);
      try {
        if (typeof caseData.relatedPeople === 'string') {
          caseData.relatedPeople = JSON.parse(caseData.relatedPeople);
          Logger.log(`[addCase] 解析相關人字串成功: ${JSON.stringify(caseData.relatedPeople)}`);
        }
      } catch (e) {
        Logger.log(`[addCase] 解析相關人字串失敗: ${e.toString()}, 將使用空陣列`);
        caseData.relatedPeople = [];
      }
    }
    
    // 確保 caseData.categories 是陣列
    if (caseData.categories && !Array.isArray(caseData.categories)) {
      Logger.log(`[addCase] 類別數據不是陣列，嘗試轉換`);
      try {
        if (typeof caseData.categories === 'string') {
          caseData.categories = JSON.parse(caseData.categories);
          Logger.log(`[addCase] 解析類別字串成功: ${JSON.stringify(caseData.categories)}`);
        }
      } catch (e) {
        Logger.log(`[addCase] 解析類別字串失敗: ${e.toString()}, 將使用空陣列`);
        caseData.categories = [];
      }
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    let casesSheet = ss.getSheetByName('案件');
    let relationsSheet = ss.getSheetByName('相關人關聯');
    let categoryRelationsSheet = ss.getSheetByName('類別關聯');
    
    // 如果工作表不存在，則創建它
    if (!casesSheet || !relationsSheet || !categoryRelationsSheet) {
      initializeSheets(sheetsId);
      casesSheet = ss.getSheetByName('案件');
      relationsSheet = ss.getSheetByName('相關人關聯');
      categoryRelationsSheet = ss.getSheetByName('類別關聯');
    }
    
    // 處理檔案上傳
    let files = [];
    try {
      if (caseData.files && caseData.files.length > 0) {
        files = uploadFiles(driveFolderId, caseData.files, caseData.title);
        if (!files || files.length === 0) {
          console.error('沒有檔案成功上傳');
        }
      }
    } catch (fileError) {
      console.error('處理檔案時發生錯誤:', fileError);
      return { success: false, error: '處理檔案時發生錯誤: ' + fileError.toString() };
    }
    
    // 獲取當前資料來生成新ID
    const data = getSheetData(casesSheet);
    const newId = 'c' + (data.length + 1);
    
    // 將相關人ID轉換為JSON字串，以便存儲在G欄位
    Logger.log(`[addCase] 新增案件時取得的原始相關人數據: ${JSON.stringify(caseData.relatedPeople)}, 類型: ${typeof caseData.relatedPeople}`);
    
    // 確保相關人資料是一個陣列
    let relatedPeopleArray = [];
    
    if (Array.isArray(caseData.relatedPeople)) {
      // 正常情況，相關人數據已經是陣列
      relatedPeopleArray = caseData.relatedPeople;
      Logger.log(`[addCase] 相關人資料是陣列，包含 ${relatedPeopleArray.length} 個項目`);
    } else if (typeof caseData.relatedPeople === 'string') {
      // 檢查是否為 JSON 字串
      try {
        const parsed = JSON.parse(caseData.relatedPeople);
        Logger.log(`[addCase] 已解析相關人JSON字串: ${JSON.stringify(parsed)}`);
        if (Array.isArray(parsed)) {
          relatedPeopleArray = parsed;
          Logger.log(`[addCase] 解析後是陣列，包含 ${relatedPeopleArray.length} 個項目`);
        } else if (parsed) {
          // 如果是其他值，嘗試轉換
          relatedPeopleArray = [parsed.toString()];
          Logger.log(`[addCase] 解析後不是陣列但有值，轉換為單項陣列: ${JSON.stringify(relatedPeopleArray)}`);
        }
      } catch (e) {
        // 如果不是 JSON 字串，但仍然是字串
        Logger.log(`[addCase] 解析JSON失敗: ${e.toString()}`);
        if (caseData.relatedPeople.trim()) {
          relatedPeopleArray = [caseData.relatedPeople.trim()];
          Logger.log(`[addCase] 使用原始字串作為單項陣列: ${JSON.stringify(relatedPeopleArray)}`);
        }
      }
    } else if (caseData.relatedPeople) {
      // 如果是其他非陣列、非字串但有效的值
      relatedPeopleArray = [caseData.relatedPeople.toString()];
      Logger.log(`[addCase] 其他類型的值，轉換為字串並使用單項陣列: ${JSON.stringify(relatedPeopleArray)}`);
    }
    
    // 確保類別資料是一個陣列
    let categoriesArray = [];
    
    if (Array.isArray(caseData.categories)) {
      // 正常情況，類別數據已經是陣列
      categoriesArray = caseData.categories;
      Logger.log(`[addCase] 類別資料是陣列，包含 ${categoriesArray.length} 個項目`);
    } else if (typeof caseData.categories === 'string') {
      // 檢查是否為 JSON 字串
      try {
        const parsed = JSON.parse(caseData.categories);
        Logger.log(`[addCase] 已解析類別JSON字串: ${JSON.stringify(parsed)}`);
        if (Array.isArray(parsed)) {
          categoriesArray = parsed;
          Logger.log(`[addCase] 解析後是陣列，包含 ${categoriesArray.length} 個項目`);
        } else if (parsed) {
          // 如果是其他值，嘗試轉換
          categoriesArray = [parsed.toString()];
          Logger.log(`[addCase] 解析後不是陣列但有值，轉換為單項陣列: ${JSON.stringify(categoriesArray)}`);
        }
      } catch (e) {
        // 如果不是 JSON 字串，但仍然是字串
        Logger.log(`[addCase] 解析JSON失敗: ${e.toString()}`);
        if (caseData.categories && caseData.categories.trim()) {
          categoriesArray = [caseData.categories.trim()];
          Logger.log(`[addCase] 使用原始字串作為單項陣列: ${JSON.stringify(categoriesArray)}`);
        }
      }
    } else if (caseData.categories) {
      // 如果是其他非陣列、非字串但有效的值
      categoriesArray = [caseData.categories.toString()];
      Logger.log(`[addCase] 其他類型的值，轉換為字串並使用單項陣列: ${JSON.stringify(categoriesArray)}`);
    }
    
    Logger.log(`[addCase] 處理後的相關人陣列: ${JSON.stringify(relatedPeopleArray)}, 長度: ${relatedPeopleArray.length}`);
    Logger.log(`[addCase] 處理後的類別陣列: ${JSON.stringify(categoriesArray)}, 長度: ${categoriesArray.length}`);
    
    // 檢查每個相關人ID是否都是有效的
    const validatedRelatedPeople = relatedPeopleArray.filter(id => id && typeof id === 'string' && id.trim() !== '');
    Logger.log(`[addCase] 驗證後的相關人陣列: ${JSON.stringify(validatedRelatedPeople)}, 長度: ${validatedRelatedPeople.length}`);
    
    // 檢查每個類別ID是否都是有效的
    const validatedCategories = categoriesArray.filter(id => id && typeof id === 'string' && id.trim() !== '');
    Logger.log(`[addCase] 驗證後的類別陣列: ${JSON.stringify(validatedCategories)}, 長度: ${validatedCategories.length}`);
    
    const relatedPeopleJson = JSON.stringify(validatedRelatedPeople);
    Logger.log(`[addCase] 準備儲存的相關人JSON字串: "${relatedPeopleJson}", 長度: ${relatedPeopleJson.length}`);
    
    const categoriesJson = JSON.stringify(validatedCategories);
    Logger.log(`[addCase] 準備儲存的類別JSON字串: "${categoriesJson}", 長度: ${categoriesJson.length}`);
    
    // 準備要寫入的案件資料
    const rowData = [
      newId,
      caseData.title,
      caseData.content,
      caseData.date,
      JSON.stringify(files),
      'false',
      relatedPeopleJson, // 將相關人資訊以JSON字串格式寫入G欄位
      categoriesJson, // 將類別資訊以JSON字串格式寫入H欄位
      'false' // 將 isFavorite 欄位設置為 false
    ];
    
    Logger.log(`準備新增案件資料 (ID: ${newId}): ${JSON.stringify(rowData)}`);
    
    // 添加新案件
    try {
      casesSheet.appendRow(rowData);
      Logger.log(`成功新增案件資料 (ID: ${newId})`);
      
      // 處理相關人資料
      if (validatedRelatedPeople.length > 0) {
        // 獲取關聯表中的資料以生成新ID
        const relationsData = getSheetData(relationsSheet);
        let relationIdCounter = relationsData.length + 1;
        
        // 為每個相關人添加一條關聯記錄
        validatedRelatedPeople.forEach(personId => {
          const relationId = 'rel' + relationIdCounter++;
          relationsSheet.appendRow([
            relationId,
            newId,
            '案件',
            personId
          ]);
          Logger.log(`成功新增相關人關聯記錄 (ID: ${relationId}, 案件ID: ${newId}, 相關人ID: ${personId})`);
        });
      }
      
      // 處理類別資料
      if (validatedCategories.length > 0) {
        // 獲取類別關聯表中的資料以生成新ID
        const categoryRelationsData = getSheetData(categoryRelationsSheet);
        let categoryRelationIdCounter = categoryRelationsData.length + 1;
        
        // 為每個類別添加一條關聯記錄
        validatedCategories.forEach(categoryId => {
          const relationId = 'catrel' + categoryRelationIdCounter++;
          categoryRelationsSheet.appendRow([
            relationId,
            newId,
            '案件',
            categoryId
          ]);
          Logger.log(`成功新增類別關聯記錄 (ID: ${relationId}, 案件ID: ${newId}, 類別ID: ${categoryId})`);
        });
      }
    } catch (appendError) {
      console.error('添加資料到工作表時發生錯誤:', appendError);
      return { success: false, error: '添加資料到工作表時發生錯誤: ' + appendError.toString() };
    }
    
    return { 
      success: true, 
      caseId: newId 
    };
  } catch (error) {
    console.error('新增案件時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 新增進度
function addProgress(sheetsId, driveFolderId, progressData) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!driveFolderId) {
      throw new Error('未提供 Google Drive 資料夾 ID');
    }
    
    Logger.log(`[addProgress] 開始處理新增進度請求，案件ID: ${progressData.caseId}`);
    Logger.log(`[addProgress] 收到的相關人資料: ${JSON.stringify(progressData.relatedPeople)}, 類型: ${typeof progressData.relatedPeople}`);
    
    // 確保相關人資料是陣列
    if (progressData.relatedPeople && !Array.isArray(progressData.relatedPeople)) {
      try {
        progressData.relatedPeople = JSON.parse(progressData.relatedPeople);
        Logger.log(`[addProgress] 解析相關人JSON成功: ${JSON.stringify(progressData.relatedPeople)}`);
      } catch (e) {
        Logger.log(`[addProgress] 解析相關人JSON失敗: ${e.toString()}`);
        progressData.relatedPeople = [];
      }
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    let progressSheet = ss.getSheetByName('進度');
    let relationsSheet = ss.getSheetByName('相關人關聯');
    
    // 如果工作表不存在，則創建它
    if (!progressSheet || !relationsSheet) {
      initializeSheets(sheetsId);
      progressSheet = ss.getSheetByName('進度');
      relationsSheet = ss.getSheetByName('相關人關聯');
    }
    
    // 處理檔案上傳
    let files = [];
    try {
      if (progressData.files && progressData.files.length > 0) {
        files = uploadFiles(driveFolderId, progressData.files, `進度-${progressData.caseId}`);
        if (!files || files.length === 0) {
          console.error('沒有進度檔案成功上傳');
        }
      }
    } catch (fileError) {
      console.error('處理進度檔案時發生錯誤:', fileError);
      return { success: false, error: '處理進度檔案時發生錯誤: ' + fileError.toString() };
    }
    
    // 獲取當前資料來生成新ID
    const data = getSheetData(progressSheet);
    const newId = 'pr' + (data.length + 1);
    
    // 準備要寫入的進度資料
    const rowData = [
      newId,
      progressData.caseId,
      progressData.date,
      progressData.content,
      JSON.stringify(files)
    ];
    
    Logger.log(`準備新增進度資料 (ID: ${newId}, CaseID: ${progressData.caseId}): ${JSON.stringify(rowData)}`);
    
    // 添加新進度
    try {
      progressSheet.appendRow(rowData);
      Logger.log(`成功新增進度資料 (ID: ${newId})`);
      
      // 處理相關人資料
      if (progressData.relatedPeople && progressData.relatedPeople.length > 0) {
        // 獲取關聯表中的資料以生成新ID
        const relationsData = getSheetData(relationsSheet);
        let relationIdCounter = relationsData.length + 1;
        
        // 為每個相關人添加一條關聯記錄
        progressData.relatedPeople.forEach(personId => {
          const relationId = 'rel' + relationIdCounter++;
          relationsSheet.appendRow([
            relationId,
            newId,
            '進度',
            personId
          ]);
          Logger.log(`成功新增相關人關聯記錄 (ID: ${relationId}, 進度ID: ${newId}, 相關人ID: ${personId})`);
        });
      }

      // 準備回傳給前端的完整進度物件
      const newProgressItem = {
        id: newId,
        caseId: progressData.caseId,
        date: progressData.date,
        relatedPeople: progressData.relatedPeople, // 使用傳入的相關人 ID 陣列
        content: progressData.content,
        files: files // 使用上傳成功後的文件資訊陣列
      };
      Logger.log(`準備回傳的新進度物件: ${JSON.stringify(newProgressItem)}`);

      // 返回成功狀態和新建立的進度物件
      return JSON.stringify({
        success: true,
        progress: newProgressItem // 回傳完整物件
      });

    } catch (appendError) {
      console.error('添加進度資料到工作表時發生錯誤:', appendError);
      return JSON.stringify({ success: false, error: '添加進度資料到工作表時發生錯誤: ' + appendError.toString() });
    }

  } catch (error) {
    console.error('新增進度時發生錯誤:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// 標記案件為已完成
function markCaseAsCompleted(sheetsId, caseId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    
    if (!casesSheet) {
      throw new Error('找不到案件工作表');
    }
    
    // 尋找案件行
    const data = getSheetData(casesSheet);
    const caseIndex = data.findIndex(row => row.id === caseId);
    
    if (caseIndex < 0) {
      throw new Error('找不到指定的案件');
    }
    
    const caseData = data[caseIndex];
    
    // 檢查案件是否已被刪除
    if (caseData.isDeleted === true || caseData.isDeleted === 'true' || String(caseData.isDeleted).toUpperCase() === 'TRUE') {
      return { success: false, error: '無法標記已刪除的案件為完成狀態' };
    }
    
    // 更新完成狀態
    const rowIndex = caseIndex + 2; // +2 是因為標題行和 0-index
    casesSheet.getRange(rowIndex, 7).setValue('true');
    
    return { success: true };
  } catch (error) {
    console.error('標記案件為已完成時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 取消案件完成標記
function cancelCaseCompletion(sheetsId, caseId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    
    if (!casesSheet) {
      throw new Error('找不到案件工作表');
    }
    
    // 尋找案件行
    const data = getSheetData(casesSheet);
    const caseIndex = data.findIndex(row => row.id === caseId);
    
    if (caseIndex < 0) {
      throw new Error('找不到指定的案件');
    }
    
    const caseData = data[caseIndex];
    
    // 檢查案件是否已被刪除
    if (caseData.isDeleted === true || caseData.isDeleted === 'true' || String(caseData.isDeleted).toUpperCase() === 'TRUE') {
      return { success: false, error: '無法取消已刪除案件的完成狀態' };
    }
    
    // 更新完成狀態為 false
    const rowIndex = caseIndex + 2; // +2 是因為標題行和 0-index
    casesSheet.getRange(rowIndex, 7).setValue('false');
    
    return { success: true };
  } catch (error) {
    console.error('取消案件完成狀態時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 上傳檔案到 Google Drive
function uploadFiles(driveFolderId, files, folderName) {
  try {
    // 獲取目標資料夾
    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(driveFolderId);
    } catch (folderError) {
      console.error('獲取目標資料夾時發生錯誤:', folderError);
      throw new Error('無法找到指定的 Google Drive 資料夾: ' + folderError.toString());
    }
    
    // 檢查是否存在案件資料夾
    let caseFolder;
    try {
      const caseFolders = targetFolder.getFoldersByName('案件進度追蹤系統');
      if (caseFolders.hasNext()) {
        caseFolder = caseFolders.next();
      } else {
        caseFolder = targetFolder.createFolder('案件進度追蹤系統');
      }
    } catch (caseFolderError) {
      console.error('建立或獲取案件資料夾時發生錯誤:', caseFolderError);
      throw new Error('建立或獲取案件資料夾時發生錯誤: ' + caseFolderError.toString());
    }
    
    // 檢查是否存在特定案件/進度的資料夾
    let specificFolder;
    try {
      const specificFolders = caseFolder.getFoldersByName(folderName);
      if (specificFolders.hasNext()) {
        specificFolder = specificFolders.next();
      } else {
        specificFolder = caseFolder.createFolder(folderName);
      }
    } catch (specificFolderError) {
      console.error('建立或獲取特定案件資料夾時發生錯誤:', specificFolderError);
      throw new Error('建立或獲取特定案件資料夾時發生錯誤: ' + specificFolderError.toString());
    }
    
    // 上傳檔案
    const uploadedFiles = [];
    const fileErrors = [];
    
    files.forEach((fileData, index) => {
      try {
        // 解碼 base64 資料
        const base64Data = fileData.content.split(',')[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileData.mimeType, fileData.name);
        
        // 儲存到 Drive
        const file = specificFolder.createFile(blob);
        
        // 設置檔案為任何人都可以查看和下載
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // 獲取檔案 ID
        const fileId = file.getId();
        
        // 使用 Google Drive 的預覽 URL
        const drivePreviewUrl = `https://drive.google.com/file/d/${fileId}/preview`;
        
        Logger.log(`檔案 ${fileData.name} (${fileData.type}) 上傳成功，ID: ${fileId}`);
        Logger.log(`為檔案創建的 Drive 預覽 URL: ${drivePreviewUrl}`);
        
        // 添加到結果列表 - 使用 Drive 預覽 URL
        uploadedFiles.push({
          id: fileId,
          name: fileData.name,
          type: fileData.type,
          previewUrl: drivePreviewUrl // 修改鍵名為 previewUrl
        });
      } catch (error) {
        console.error(`上傳第 ${index + 1} 個檔案時發生錯誤:`, error);
        fileErrors.push(`檔案 "${fileData.name}": ${error.toString()}`);
      }
    });
    
    if (fileErrors.length > 0 && fileErrors.length === files.length) {
      throw new Error('所有檔案上傳失敗: ' + fileErrors.join('; '));
    }
    
    if (fileErrors.length > 0) {
      console.warn('部分檔案上傳失敗:', fileErrors.join('; '));
    }
    
    return uploadedFiles;
  } catch (error) {
    console.error('處理檔案資料夾時發生錯誤:', error);
    throw error; // 將錯誤向上傳遞
  }
}

// 從工作表獲取資料並轉換為對象陣列
function getSheetData(sheet) {
  try {
    if (!sheet) {
      throw new Error('無效的工作表');
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      console.log(`工作表 "${sheet.getName()}" 只有標題行或無資料`);
      return [];
    }
    
    const headers = data[0];
    
    // 驗證標題行
    if (headers.length === 0) {
      throw new Error(`工作表 "${sheet.getName()}" 標題行為空`);
    }
    
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      try {
        const row = data[i];
        const obj = {};
        
        // 跳過全空的行
        let hasValue = false;
        for (let j = 0; j < row.length; j++) {
          if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
            hasValue = true;
            break;
          }
        }
        
        if (!hasValue) {
          // console.log(`跳過空行 - 行 ${i+1}`); // 暫時註解掉，減少日誌量
          continue;
        }
        
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) { // 確保標題不為空
            // === 新增偵錯：記錄原始值和賦值 ===
            const headerName = headers[j];
            const rawValue = row[j];
            if (sheet.getName() === '案件' && headerName === 'files') {
              Logger.log(`[getSheetData - 案件 - 行 ${i+1}] 處理欄位 '${headerName}' (j=${j})，原始值: ${JSON.stringify(rawValue)}, 類型: ${typeof rawValue}`);
            }
            // === 結束新增偵錯 ===
            
            // 特殊處理類別工作表的 order 欄位，確保其為數值類型
            if (sheet.getName() === '類別' && headerName === 'order') {
              obj[headerName] = rawValue ? Number(rawValue) : 9999;
              Logger.log(`[getSheetData - 類別 - 行 ${i+1}] 處理 order 欄位，原始值: ${rawValue}, 轉換後: ${obj[headerName]}`);
            } else {
              obj[headerName] = rawValue;
            }
            
            // === 新增偵錯：記錄賦值後的值 ===
            if (sheet.getName() === '案件' && headerName === 'files') {
                Logger.log(`[getSheetData - 案件 - 行 ${i+1}] 賦值給 obj.${headerName} 後的值: ${JSON.stringify(obj[headerName])}, 類型: ${typeof obj[headerName]}`);
            }
            // === 結束新增偵錯 ===
          }
        }
        
        // 強制轉換特定欄位為字串 (如果需要)
        if (sheet.getName() === '案件') {
          if (obj.hasOwnProperty('title') && typeof obj.title !== 'string') {
            obj.title = obj.title !== null && obj.title !== undefined ? String(obj.title) : '';
            // Logger.log(`已將案件 ${obj.id} 的 title 強制轉換為字串`); // 暫時註解掉
          }
          if (obj.hasOwnProperty('content') && typeof obj.content !== 'string') {
            obj.content = obj.content !== null && obj.content !== undefined ? String(obj.content) : '';
            // Logger.log(`已將案件 ${obj.id} 的 content 強制轉換為字串`); // 暫時註解掉
          }
          // === 新增偵錯：檢查物件轉換後的 files 值 ===
          if (obj.hasOwnProperty('files')) {
              Logger.log(`[getSheetData - 案件 - 行 ${i+1}] 完成該行物件轉換後，obj.files 的值: ${JSON.stringify(obj.files)}, 類型: ${typeof obj.files}`);
          }
          // === 結束新增偵錯 ===
        } else if (sheet.getName() === '進度') {
          if (obj.hasOwnProperty('content') && typeof obj.content !== 'string') {
            obj.content = obj.content !== null && obj.content !== undefined ? String(obj.content) : '';
            // Logger.log(`已將進度 ${obj.id} (案件 ${obj.caseId}) 的 content 強制轉換為字串`); // 暫時註解掉
          }
          // 可以為進度也加上 files 的檢查
        }
        
        result.push(obj);
      } catch (rowError) {
        console.error(`處理工作表 "${sheet.getName()}" 第 ${i+1} 行時發生錯誤:`, rowError);
        // 繼續處理下一行
      }
    }
    
    // Logger.log(`從工作表 "${sheet.getName()}" 轉換後的資料 (前${Math.min(5, result.length)}筆): ${JSON.stringify(result.slice(0, 5))}`); // 暫時註解掉以專注於 files 欄位
    return result;
  } catch (error) {
    console.error(`從工作表 "${sheet ? sheet.getName() : '未知'}" 獲取資料時發生錯誤:`, error);
    throw error; // 向上傳遞錯誤
  }
}

// 獲取特定案件詳情
function getCaseDetails(sheetsId, caseId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    const progressSheet = ss.getSheetByName('進度');
    const relationsSheet = ss.getSheetByName('相關人關聯');
    const peopleSheet = ss.getSheetByName('相關人');
    
    if (!casesSheet || !progressSheet || !relationsSheet) {
      throw new Error('找不到必要的工作表');
    }
    
    // 讀取案件資料
    const casesData = getSheetData(casesSheet);
    const progressData = getSheetData(progressSheet);
    const relationsData = getSheetData(relationsSheet);
    const peopleData = getSheetData(peopleSheet);
    
    // 找到指定的案件
    const caseItem = casesData.find(c => c.id === caseId);
    
    if (!caseItem) {
      return JSON.stringify({ success: false, error: '找不到指定的案件' });
    }
    
    // 新增偵錯記錄
    Logger.log(`[getCaseDetails] 案件ID: ${caseId}, 原始案件資料:`, JSON.stringify(caseItem));
    
    // 從關聯表中獲取該案件的相關人ID
    const caseRelations = relationsData.filter(relation => 
      relation.objectId === caseId && relation.objectType === '案件');
    const relatedPeopleFromRelations = caseRelations.map(relation => relation.personId);
    
    // 檢查案件資料中是否有直接存儲的相關人資訊
    let relatedPeople = relatedPeopleFromRelations;
    if (caseItem.relatedPeople && typeof caseItem.relatedPeople === 'string' && caseItem.relatedPeople.trim() !== '') {
      try {
        const parsedRelatedPeople = JSON.parse(caseItem.relatedPeople);
        Logger.log(`[getCaseDetails] 案件 ${caseId} G欄位相關人資訊: ${caseItem.relatedPeople}, 解析後:`, JSON.stringify(parsedRelatedPeople));
        if (Array.isArray(parsedRelatedPeople) && parsedRelatedPeople.length > 0) {
          Logger.log(`[案件詳情 - ID: ${caseId}] 從 G 欄位讀取到相關人資訊: ${caseItem.relatedPeople}`);
          relatedPeople = parsedRelatedPeople; // 優先使用 G 欄位的相關人資訊
        } else {
          Logger.log(`[案件詳情 - ID: ${caseId}] G 欄位相關人資訊為空陣列或非陣列`);
        }
      } catch (jsonError) {
        Logger.log(`[案件詳情 - ID: ${caseId}] 解析 G 欄位相關人資訊時發生錯誤: ${jsonError}`);
      }
    } else {
      Logger.log(`[案件詳情 - ID: ${caseId}] G 欄位相關人資訊不存在或為空`);
    }
    
    Logger.log(`[getCaseDetails] 最終使用的相關人列表: ${JSON.stringify(relatedPeople)}, 來源: ${relatedPeople === relatedPeopleFromRelations ? '關聯表' : 'G欄位'}`);
    
    // 獲取案件的所有進度
    const progressItems = [];
    progressData.forEach(progressRow => {
      if (progressRow.caseId === caseId) {
        // 從關聯表中獲取該進度的相關人ID
        const progressRelations = relationsData.filter(relation => 
          relation.objectId === progressRow.id && relation.objectType === '進度');
        const progressRelatedPeople = progressRelations.map(relation => relation.personId);
        
        progressItems.push({
          id: progressRow.id,
          date: progressRow.date,
          relatedPeople: progressRelatedPeople,
          content: progressRow.content,
          files: progressRow.files ? JSON.parse(progressRow.files) : []
        });
      }
    });
    
    // 格式化案件資料
    const result = {
      id: caseItem.id,
      title: caseItem.title,
      content: caseItem.content,
      date: caseItem.date,
      relatedPeople: relatedPeople,
      files: caseItem.files ? JSON.parse(caseItem.files) : [],
      progress: progressItems,
      completed: caseItem.completed === 'true'
    };
    
    const finalResult = { success: true, caseDetails: result };
    Logger.log('案件詳情查詢結果 (序列化前): ' + JSON.stringify(finalResult));
    
    // 將結果序列化為 JSON 字串返回，與 getAllCases 保持一致
    return JSON.stringify(finalResult);
  } catch (error) {
    console.error('獲取案件詳情時發生錯誤:', error);
    const errorResult = { success: false, error: error.toString() };
    Logger.log('案件詳情查詢錯誤 (序列化前): ' + JSON.stringify(errorResult));
    
    // 將錯誤結果也序列化返回
    return JSON.stringify(errorResult);
  }
}

// 應用日期篩選獲取案件列表
function filterCasesByDate(sheetsId, filterType, startDate, endDate, month) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    const relationsSheet = ss.getSheetByName('相關人關聯');
    
    if (!casesSheet || !relationsSheet) {
      throw new Error('找不到必要的工作表');
    }
    
    // 讀取案件資料
    const casesData = getSheetData(casesSheet);
    const relationsData = getSheetData(relationsSheet);
    
    // 根據篩選類型過濾資料
    let filteredCases = [];
    
    if (filterType === 'range') {
      const start = new Date(startDate);
      const end = new Date(endDate);
      end.setHours(23, 59, 59); // 設置為當天最後一刻
      
      filteredCases = casesData.filter(caseItem => {
        const caseDate = new Date(caseItem.date);
        return caseDate >= start && caseDate <= end && caseItem.completed !== 'true';
      });
    } else if (filterType === 'month') {
      const [year, monthStr] = month.split('-');
      const monthIndex = parseInt(monthStr) - 1; // 月份從 0 開始
      
      filteredCases = casesData.filter(caseItem => {
        const caseDate = new Date(caseItem.date);
        return caseDate.getFullYear() === parseInt(year) && 
               caseDate.getMonth() === monthIndex && 
               caseItem.completed !== 'true';
      });
    }
    
    // 格式化案件資料
    const result = filteredCases.map(caseRow => {
      // 從關聯表中獲取該案件的相關人ID
      const caseRelations = relationsData.filter(relation => 
        relation.objectId === caseRow.id && relation.objectType === '案件');
      const relatedPeopleFromRelations = caseRelations.map(relation => relation.personId);
      
      // 檢查案件資料中是否有直接存儲的相關人資訊
      let relatedPeople = relatedPeopleFromRelations;
      if (caseRow.relatedPeople && typeof caseRow.relatedPeople === 'string' && caseRow.relatedPeople.trim() !== '') {
        try {
          const parsedRelatedPeople = JSON.parse(caseRow.relatedPeople);
          if (Array.isArray(parsedRelatedPeople) && parsedRelatedPeople.length > 0) {
            // 優先使用 G 欄位的相關人資訊
            relatedPeople = parsedRelatedPeople;
          }
        } catch (jsonError) {
          // 發生錯誤時使用關聯表中的相關人資訊
          console.error(`篩選案件 - [案件 ID: ${caseRow.id}] 解析 G 欄位相關人資訊時發生錯誤:`, jsonError);
        }
      }
      
      return {
        id: caseRow.id,
        title: caseRow.title,
        content: caseRow.content,
        date: caseRow.date,
        relatedPeople: relatedPeople,
        files: caseRow.files ? JSON.parse(caseRow.files) : [],
        completed: caseRow.completed === 'true'
      };
    });
    
    return { success: true, cases: result };
  } catch (error) {
    console.error('篩選案件時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 應用日期篩選獲取已完成案件列表
function filterCompletedCasesByDate(sheetsId, filterType, startDate, endDate, month) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    const relationsSheet = ss.getSheetByName('相關人關聯');
    
    if (!casesSheet || !relationsSheet) {
      throw new Error('找不到必要的工作表');
    }
    
    // 讀取案件資料
    const casesData = getSheetData(casesSheet);
    const relationsData = getSheetData(relationsSheet);
    
    // 根據篩選類型過濾資料
    let filteredCases = [];
    
    if (filterType === 'range') {
      const start = new Date(startDate);
      const end = new Date(endDate);
      end.setHours(23, 59, 59); // 設置為當天最後一刻
      
      filteredCases = casesData.filter(caseItem => {
        const caseDate = new Date(caseItem.date);
        return caseDate >= start && caseDate <= end && caseItem.completed === 'true';
      });
    } else if (filterType === 'month') {
      const [year, monthStr] = month.split('-');
      const monthIndex = parseInt(monthStr) - 1; // 月份從 0 開始
      
      filteredCases = casesData.filter(caseItem => {
        const caseDate = new Date(caseItem.date);
        return caseDate.getFullYear() === parseInt(year) && 
               caseDate.getMonth() === monthIndex && 
               caseItem.completed === 'true';
      });
    }
    
    // 格式化案件資料
    const result = filteredCases.map(caseRow => {
      // 從關聯表中獲取該案件的相關人ID
      const caseRelations = relationsData.filter(relation => 
        relation.objectId === caseRow.id && relation.objectType === '案件');
      const relatedPeopleFromRelations = caseRelations.map(relation => relation.personId);
      
      // 檢查案件資料中是否有直接存儲的相關人資訊
      let relatedPeople = relatedPeopleFromRelations;
      if (caseRow.relatedPeople && typeof caseRow.relatedPeople === 'string' && caseRow.relatedPeople.trim() !== '') {
        try {
          const parsedRelatedPeople = JSON.parse(caseRow.relatedPeople);
          if (Array.isArray(parsedRelatedPeople) && parsedRelatedPeople.length > 0) {
            // 優先使用 G 欄位的相關人資訊
            relatedPeople = parsedRelatedPeople;
          }
        } catch (jsonError) {
          // 發生錯誤時使用關聯表中的相關人資訊
          console.error(`篩選已完成案件 - [案件 ID: ${caseRow.id}] 解析 G 欄位相關人資訊時發生錯誤:`, jsonError);
        }
      }
      
      return {
        id: caseRow.id,
        title: caseRow.title,
        content: caseRow.content,
        date: caseRow.date,
        relatedPeople: relatedPeople,
        files: caseRow.files ? JSON.parse(caseRow.files) : [],
        completed: caseRow.completed === 'true'
      };
    });
    
    return { success: true, cases: result };
  } catch (error) {
    console.error('篩選已完成案件時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 檢查工作表結構並初始化
function initializeSheets(sheetsId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    
    // 檢查並創建必要的工作表
    const requiredSheets = ['案件', '進度', '相關人關聯', '相關人', '類別', '類別關聯'];
    const existingSheets = ss.getSheets().map(sheet => sheet.getName());
    
    requiredSheets.forEach(sheetName => {
      if (!existingSheets.includes(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        
        // 設置標題行
        if (sheetName === '案件') {
          sheet.appendRow(['id', 'title', 'content', 'date', 'files', 'completed', 'relatedPeople', 'categories', 'isFavorite']);
        } else if (sheetName === '進度') {
          sheet.appendRow(['id', 'caseId', 'date', 'content', 'files']);
        } else if (sheetName === '相關人關聯') {
          sheet.appendRow(['id', 'objectId', 'objectType', 'personId']);  // objectType可以是"案件"或"進度"
        } else if (sheetName === '相關人') {
          sheet.appendRow(['id', 'name', 'color', 'createDate', 'isActive']);
        } else if (sheetName === '類別') {
          sheet.appendRow(['id', 'name', 'color', 'createDate', 'isActive', 'order']);
        } else if (sheetName === '類別關聯') {
          sheet.appendRow(['id', 'objectId', 'objectType', 'categoryId']);  // objectType可以是"案件"或"進度"
        }
      }
    });
    
    // 檢查並修復現有工作表的標題行
    const fixedSheets = checkAndFixSheetHeaders(ss);
    
    // 返回結果，包括修復的工作表信息
    return { 
      success: true, 
      message: '已初始化工作表', 
      fixedSheets: fixedSheets 
    };
  } catch (error) {
    console.error('初始化工作表時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

// 檢查並修復工作表標題
function checkAndFixSheetHeaders(spreadsheet) {
  try {
    Logger.log('開始檢查和修復工作表標題');
    const fixedSheets = [];
    
    // 確定每個工作表應該有的標題
    const expectedHeaders = {
      '案件': ['id', 'title', 'content', 'date', 'files', 'completed', 'relatedPeople', 'categories', 'isFavorite'],
      '進度': ['id', 'caseId', 'date', 'content', 'files'],
      '相關人關聯': ['id', 'objectId', 'objectType', 'personId'],
      '相關人': ['id', 'name', 'color', 'createDate', 'isActive'],
      '類別': ['id', 'name', 'color', 'createDate', 'isActive', 'order'],
      '類別關聯': ['id', 'objectId', 'objectType', 'categoryId']
    };
    
    // 檢查每個工作表
    Object.keys(expectedHeaders).forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) return; // 如果工作表不存在，則略過
      
      // 檢查工作表是否有數據
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      if (lastRow === 0) {
        // 工作表為空，直接添加標題行
        sheet.appendRow(expectedHeaders[sheetName]);
        fixedSheets.push({
          name: sheetName,
          action: '添加標題行',
          headers: expectedHeaders[sheetName].join(', ')
        });
      } else {
        // 檢查第一行是否為標題行
        const firstRow = sheet.getRange(1, 1, 1, expectedHeaders[sheetName].length).getValues()[0];
        
        // 檢查每個列是否有標題
        let needsFix = false;
        for (let i = 0; i < expectedHeaders[sheetName].length; i++) {
          if (i >= firstRow.length || !firstRow[i]) {
            needsFix = true;
            break;
          }
        }
        
        // 如果需要修復，檢查第一行是否包含數據而非標題
        if (needsFix) {
          // 判斷第一行是否為數據而非標題（如果第一個單元格是數字或GUID格式，則很可能是數據）
          const isFirstRowData = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(String(firstRow[0])) || 
                                !isNaN(firstRow[0]);
          
          if (isFirstRowData) {
            // 第一行是數據，在前面插入標題行
            sheet.insertRowBefore(1);
            sheet.getRange(1, 1, 1, expectedHeaders[sheetName].length).setValues([expectedHeaders[sheetName]]);
            fixedSheets.push({
              name: sheetName,
              action: '在數據前插入標題行',
              headers: expectedHeaders[sheetName].join(', ')
            });
          } else {
            // 第一行可能是不完整的標題，更新它
            sheet.getRange(1, 1, 1, expectedHeaders[sheetName].length).setValues([expectedHeaders[sheetName]]);
            fixedSheets.push({
              name: sheetName,
              action: '更新不完整的標題行',
              headers: expectedHeaders[sheetName].join(', ')
            });
          }
        }
      }
    });
    
    Logger.log(`完成檢查和修復工作表標題，共修復 ${fixedSheets.length} 個工作表`);
    return fixedSheets;
  } catch (error) {
    Logger.log(`檢查和修復工作表標題時發生錯誤: ${error}`);
    return [];
  }
}

// 獲取系統日誌
function getDebugLogs() {
  try {
    // 取得 Apps Script 日誌
    const logs = Logger.getLog();
    
    // 不再清空日誌，這樣可以累積保留，避免顯示"無日誌內容"
    // Logger.clear(); // 移除這行
    
    // 如果日誌為空，則嘗試添加一些診斷信息
    if (!logs || logs.trim() === '') {
      Logger.log('獲取日誌時間: ' + new Date().toISOString());
      Logger.log('診斷信息: 系統日誌中沒有內容。這可能是因為:');
      Logger.log('1. 尚未有任何操作產生日誌');
      Logger.log('2. 日誌可能已被自動清除');
      
      // 嘗試記錄一些基本的系統信息
      try {
        const scriptProperties = PropertiesService.getScriptProperties().getProperties();
        const keys = Object.keys(scriptProperties);
        Logger.log('系統已配置屬性數: ' + keys.length);
        
        const ss = SpreadsheetApp.getActive();
        if (ss) {
          Logger.log('當前活躍試算表: ' + ss.getName() + ' (ID: ' + ss.getId() + ')');
          Logger.log('工作表數量: ' + ss.getSheets().length);
          Logger.log('工作表名稱: ' + ss.getSheets().map(sheet => sheet.getName()).join(', '));
        } else {
          Logger.log('當前沒有活躍試算表');
        }
      } catch (e) {
        Logger.log('嘗試獲取系統信息時發生錯誤: ' + e.toString());
      }
    }
    
    // 再次獲取日誌，包含剛剛添加的診斷信息
    const updatedLogs = Logger.getLog();
    
    return {
      success: true,
      logs: updatedLogs || '真的沒有任何日誌內容，請嘗試執行一些操作後再試',
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('獲取日誌時發生錯誤:', error);
    return {
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    };
  }
}

// 測試連接試算表
function testSheetsConnection(sheetsId) {
  try {
    Logger.log('測試試算表連接，ID: ' + sheetsId);
    
    if (!sheetsId) {
      Logger.log('錯誤：未提供 Google Sheets ID');
      return {
        success: false,
        error: '未提供 Google Sheets ID',
        details: '請確保已提供有效的 Google Sheets ID'
      };
    }
    
    let ss;
    try {
      Logger.log('嘗試開啟試算表，ID: ' + sheetsId);
      ss = SpreadsheetApp.openById(sheetsId);
      Logger.log('成功開啟試算表: ' + ss.getName());
      
      // 取得試算表基本信息
      const details = {
        name: ss.getName(),
        url: ss.getUrl(),
        sheets: ss.getSheets().map(sheet => sheet.getName()),
        owner: Session.getEffectiveUser().getEmail()
      };
      
      return {
        success: true,
        message: '成功連接到試算表',
        details: details
      };
    } catch (sheetError) {
      Logger.log('打開試算表時發生錯誤: ' + JSON.stringify(sheetError));
      Logger.log('錯誤堆疊: ' + (sheetError.stack || '無堆疊信息'));
      
      return {
        success: false,
        error: '無法開啟試算表',
        errorType: sheetError.name || '未知錯誤類型',
        errorMessage: sheetError.toString(),
        errorStack: sheetError.stack || '無堆疊信息',
        possibleCauses: [
          '試算表 ID 不正確',
          '沒有足夠的權限訪問試算表',
          '試算表已被刪除',
          'Google Apps Script 服務中斷'
        ]
      };
    }
  } catch (error) {
    Logger.log('測試連接時發生未知錯誤: ' + error);
    return {
      success: false,
      error: '測試連接時發生未知錯誤',
      details: error.toString()
    };
  }
}

// 獲取相關人列表
function getPeople(sheetsId) {
  try {
    Logger.log('獲取相關人列表，使用的sheetsId: ' + sheetsId);
    
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const peopleSheet = ss.getSheetByName('相關人');
    
    if (!peopleSheet) {
      // 如果工作表不存在，則初始化
      initializeSheets(sheetsId);
      return JSON.stringify({ 
        success: true, 
        people: []
      });
    }
    
    // 讀取相關人資料
    const peopleData = getSheetData(peopleSheet);
    
    // 格式化相關人資料
    const result = peopleData.map(person => {
      return {
        id: person.id,
        name: person.name,
        color: person.color || generateRandomColor(),
        createDate: person.createDate || new Date().toISOString(),
        isActive: !(person.isActive === 'false' || person.isActive === false)
      };
    });
    
    return JSON.stringify({ 
      success: true, 
      people: result 
    });
  } catch (error) {
    console.error('獲取相關人列表時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 新增相關人
function addPerson(sheetsId, personData) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!personData || !personData.name) {
      throw new Error('未提供相關人姓名');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    let peopleSheet = ss.getSheetByName('相關人');
    
    // 如果工作表不存在，則創建它
    if (!peopleSheet) {
      initializeSheets(sheetsId);
      peopleSheet = ss.getSheetByName('相關人');
    }
    
    // 獲取當前資料來生成新ID
    const data = getSheetData(peopleSheet);
    
    // 先檢查是否有同名但非活躍的相關人
    const inactivePerson = data.find(p => p.name === personData.name && p.isActive === 'false');
    if (inactivePerson) {
      // 重新啟用這個相關人
      const rowIndex = data.findIndex(p => p.id === inactivePerson.id) + 2;
      peopleSheet.getRange(rowIndex, 5).setValue('true');
      
      return JSON.stringify({ 
        success: true, 
        person: {
          id: inactivePerson.id,
          name: inactivePerson.name,
          color: inactivePerson.color,
          createDate: inactivePerson.createDate,
          isActive: true
        },
        message: '已恢復存在的相關人'
      });
    }
    
    // 檢查是否已存在活躍的同名相關人
    const existingPerson = data.find(p => p.name === personData.name && p.isActive !== 'false');
    if (existingPerson) {
      return JSON.stringify({ 
        success: false, 
        error: '此相關人已存在' 
      });
    }
    
    const newId = 'p' + (data.length + 1);
    const color = personData.color || generateRandomColor();
    const now = new Date().toISOString();
    
    // 準備要寫入的相關人資料
    const rowData = [
      newId,
      personData.name,
      color,
      now,
      'true'
    ];
    
    // 添加新相關人
    peopleSheet.appendRow(rowData);
    
    return JSON.stringify({ 
      success: true, 
      person: {
        id: newId,
        name: personData.name,
        color: color,
        createDate: now,
        isActive: true
      }
    });
  } catch (error) {
    console.error('新增相關人時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 編輯相關人
function updatePerson(sheetsId, personId, personData) {
  try {
    if (!sheetsId || !personId) {
      throw new Error('未提供 Google Sheets ID 或相關人ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const peopleSheet = ss.getSheetByName('相關人');
    
    if (!peopleSheet) {
      throw new Error('找不到相關人工作表');
    }
    
    // 尋找相關人行
    const data = getSheetData(peopleSheet);
    const rowIndex = data.findIndex(row => row.id === personId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的相關人');
    }
    
    // 更新名稱
    if (personData.name !== undefined) {
      peopleSheet.getRange(rowIndex, 2).setValue(personData.name);
    }
    
    // 更新顏色
    if (personData.color !== undefined) {
      peopleSheet.getRange(rowIndex, 3).setValue(personData.color);
    }
    
    // 更新活躍狀態
    if (personData.isActive !== undefined) {
      peopleSheet.getRange(rowIndex, 5).setValue(personData.isActive ? 'true' : 'false');
    }
    
    return JSON.stringify({ 
      success: true,
      person: {
        id: personId,
        name: personData.name !== undefined ? personData.name : data[rowIndex - 2].name,
        color: personData.color !== undefined ? personData.color : (data[rowIndex - 2].color || generateRandomColor()),
        createDate: data[rowIndex - 2].createDate,
        isActive: personData.isActive !== undefined ? personData.isActive : (data[rowIndex - 2].isActive !== 'false')
      }
    });
  } catch (error) {
    console.error('更新相關人時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 刪除相關人（標記為非活躍）
function removePerson(sheetsId, personId) {
  try {
    if (!sheetsId || !personId) {
      throw new Error('未提供 Google Sheets ID 或相關人ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const peopleSheet = ss.getSheetByName('相關人');
    
    if (!peopleSheet) {
      throw new Error('找不到相關人工作表');
    }
    
    // 尋找相關人行
    const data = getSheetData(peopleSheet);
    const rowIndex = data.findIndex(row => row.id === personId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的相關人');
    }
    
    // 標記為非活躍
    peopleSheet.getRange(rowIndex, 5).setValue('false');
    
    return JSON.stringify({ 
      success: true,
      message: '已將相關人標記為非活躍'
    });
  } catch (error) {
    console.error('刪除相關人時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 生成隨機顏色
function generateRandomColor() {
  const colors = [
    '#FFB6C1', '#FFA07A', '#FFDAB9', '#FFFACD', '#E0FFFF', 
    '#B0E0E6', '#D8BFD8', '#FFE4E1', '#98FB98', '#AFEEEE',
    '#7FFFD4', '#FFD700', '#FFA500', '#40E0D0', '#87CEFA'
  ];
  return colors[Math.floor(Math.random() * colors.length)];
}

// 獲取類別列表
function getCategories(sheetsId) {
  try {
    Logger.log('獲取類別列表，使用的sheetsId: ' + sheetsId);
    
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const categoriesSheet = ss.getSheetByName('類別');
    
    if (!categoriesSheet) {
      // 如果工作表不存在，則初始化
      initializeSheets(sheetsId);
      return JSON.stringify({ 
        success: true, 
        categories: []
      });
    }
    
    // 檢查是否有order欄位，沒有則添加
    const headers = categoriesSheet.getRange(1, 1, 1, categoriesSheet.getLastColumn()).getValues()[0];
    if (!headers.includes('order')) {
      updateCategoryTableStructure(sheetsId);
    }
    
    // 讀取類別資料
    const categoriesData = getSheetData(categoriesSheet);
    Logger.log('從工作表讀取到的原始類別數據: ' + JSON.stringify(categoriesData));
    
    // 記錄每個類別的原始順序值
    categoriesData.forEach(cat => {
      Logger.log(`類別 ${cat.id} (${cat.name}) 的原始 order 值: ${cat.order}, 類型: ${typeof cat.order}`);
    });
    
    // 格式化類別資料 - 確保所有屬性類型正確
    let result = categoriesData.map(category => {
      // 確保order是數字類型
      let orderValue = category.order;
      if (typeof orderValue === 'string') {
        orderValue = Number(orderValue);
        if (isNaN(orderValue)) orderValue = 9999;
        Logger.log(`類別 ${category.id} (${category.name}) 的 order 從字串 "${category.order}" 轉換為數字 ${orderValue}`);
      } else if (orderValue === undefined || orderValue === null) {
        orderValue = 9999;
        Logger.log(`類別 ${category.id} (${category.name}) 沒有 order 值，設為預設值 ${orderValue}`);
      } else {
        orderValue = Number(orderValue);
        Logger.log(`類別 ${category.id} (${category.name}) 的 order 值是 ${orderValue}, 轉換前類型: ${typeof category.order}`);
      }
      
      return {
        id: category.id,
        name: category.name,
        color: category.color || generateRandomColor(),
        createDate: category.createDate || new Date().toISOString(),
        isActive: !(category.isActive === 'false' || category.isActive === false),
        order: orderValue // 使用轉換後的數字
      };
    });
    
    // 按order欄位排序前的數據
    Logger.log('排序前的類別數據: ' + JSON.stringify(result.map(c => ({id: c.id, name: c.name, order: c.order, orderType: typeof c.order}))));
    
    // 強制確保所有 order 值都是數字
    for (let i = 0; i < result.length; i++) {
      if (typeof result[i].order !== 'number') {
        const oldValue = result[i].order;
        result[i].order = Number(result[i].order) || 9999;
        Logger.log(`強制修正：類別 ${result[i].id} (${result[i].name}) 的 order 從 ${oldValue} (${typeof oldValue}) 修正為 ${result[i].order}`);
      }
    }
    
    // 簡化排序邏輯，直接使用數字比較
    result.sort((a, b) => {
      // 確保 a.order 和 b.order 都是數字
      const orderA = Number(a.order);
      const orderB = Number(b.order);
      
      Logger.log(`排序比較: ${a.name}(order=${orderA}, ${typeof orderA}) vs ${b.name}(order=${orderB}, ${typeof orderB}), 結果: ${orderA - orderB}`);
      
      return orderA - orderB;
    });
    
    // 排序後的數據
    Logger.log('排序後的類別數據: ' + JSON.stringify(result.map(c => ({id: c.id, name: c.name, order: c.order}))));
    
    // 序列化結果前檢查排序是否生效
    let isSorted = true;
    for (let i = 1; i < result.length; i++) {
      if (Number(result[i-1].order) > Number(result[i].order)) {
        isSorted = false;
        Logger.log(`排序異常: ${result[i-1].name}(order=${result[i-1].order}) > ${result[i].name}(order=${result[i].order})`);
      }
    }
    Logger.log('類別數據是否正確排序: ' + isSorted);
    
    // 再次確認所有類別的順序值
    Logger.log('最終排序後的每個類別順序:');
    for (let i = 0; i < result.length; i++) {
      Logger.log(`[${i}] 類別 ${result[i].id} (${result[i].name}): order = ${result[i].order}, 類型: ${typeof result[i].order}`);
    }
    
    const finalResponse = JSON.stringify({ 
      success: true, 
      categories: result 
    });
    
    Logger.log('返回給前端的類別數據長度: ' + result.length);
    Logger.log('返回數據示例(全部): ' + JSON.stringify(result.map(c => ({id: c.id, name: c.name, order: c.order}))));
    
    return finalResponse;
  } catch (error) {
    console.error('獲取類別列表時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 新增類別
function addCategory(sheetsId, categoryData) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!categoryData || !categoryData.name) {
      throw new Error('未提供類別名稱');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    let categoriesSheet = ss.getSheetByName('類別');
    
    // 如果工作表不存在，則創建它
    if (!categoriesSheet) {
      initializeSheets(sheetsId);
      categoriesSheet = ss.getSheetByName('類別');
    }
    
    // 獲取當前資料來生成新ID
    const data = getSheetData(categoriesSheet);
    
    // 先檢查是否有同名但非活躍的類別
    const inactiveCategory = data.find(c => c.name === categoryData.name && c.isActive === 'false');
    if (inactiveCategory) {
      // 重新啟用這個類別
      const rowIndex = data.findIndex(c => c.id === inactiveCategory.id) + 2;
      categoriesSheet.getRange(rowIndex, 5).setValue('true');
      
      return JSON.stringify({ 
        success: true, 
        category: {
          id: inactiveCategory.id,
          name: inactiveCategory.name,
          color: inactiveCategory.color,
          createDate: inactiveCategory.createDate,
          isActive: true
        },
        message: '已恢復存在的類別'
      });
    }
    
    // 檢查是否已存在活躍的同名類別
    const existingCategory = data.find(c => c.name === categoryData.name && c.isActive !== 'false');
    if (existingCategory) {
      return JSON.stringify({ 
        success: false, 
        error: '此類別已存在' 
      });
    }
    
    const newId = 'cat' + (data.length + 1);
    const color = categoryData.color || generateRandomColor();
    const now = new Date().toISOString();
    
    // 準備要寫入的類別資料
    const rowData = [
      newId,
      categoryData.name,
      color,
      now,
      'true'
    ];
    
    // 添加新類別
    categoriesSheet.appendRow(rowData);
    
    return JSON.stringify({ 
      success: true, 
      category: {
        id: newId,
        name: categoryData.name,
        color: color,
        createDate: now,
        isActive: true
      }
    });
  } catch (error) {
    console.error('新增類別時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 編輯類別
function updateCategory(sheetsId, categoryId, categoryData) {
  try {
    if (!sheetsId || !categoryId) {
      throw new Error('未提供 Google Sheets ID 或類別ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const categoriesSheet = ss.getSheetByName('類別');
    
    if (!categoriesSheet) {
      throw new Error('找不到類別工作表');
    }
    
    // 尋找類別行
    const data = getSheetData(categoriesSheet);
    const rowIndex = data.findIndex(row => row.id === categoryId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的類別');
    }
    
    // 更新名稱
    if (categoryData.name !== undefined) {
      categoriesSheet.getRange(rowIndex, 2).setValue(categoryData.name);
    }
    
    // 更新顏色
    if (categoryData.color !== undefined) {
      categoriesSheet.getRange(rowIndex, 3).setValue(categoryData.color);
    }
    
    // 更新活躍狀態
    if (categoryData.isActive !== undefined) {
      categoriesSheet.getRange(rowIndex, 5).setValue(categoryData.isActive ? 'true' : 'false');
    }
    
    return JSON.stringify({ 
      success: true,
      category: {
        id: categoryId,
        name: categoryData.name !== undefined ? categoryData.name : data[rowIndex - 2].name,
        color: categoryData.color !== undefined ? categoryData.color : (data[rowIndex - 2].color || generateRandomColor()),
        createDate: data[rowIndex - 2].createDate,
        isActive: categoryData.isActive !== undefined ? categoryData.isActive : (data[rowIndex - 2].isActive !== 'false')
      }
    });
  } catch (error) {
    console.error('更新類別時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 刪除類別（標記為非活躍）
function removeCategory(sheetsId, categoryId) {
  try {
    if (!sheetsId || !categoryId) {
      throw new Error('未提供 Google Sheets ID 或類別ID');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const categoriesSheet = ss.getSheetByName('類別');
    
    if (!categoriesSheet) {
      throw new Error('找不到類別工作表');
    }
    
    // 尋找類別行
    const data = getSheetData(categoriesSheet);
    const rowIndex = data.findIndex(row => row.id === categoryId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的類別');
    }
    
    // 標記為非活躍
    categoriesSheet.getRange(rowIndex, 5).setValue('false');
    
    return JSON.stringify({ 
      success: true,
      message: '已將類別標記為非活躍'
    });
  } catch (error) {
    console.error('刪除類別時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 切換案件收藏狀態
function toggleCaseFavorite(sheetsId, caseId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!caseId) {
      throw new Error('未提供案件 ID');
    }
    
    Logger.log(`切換案件 ${caseId} 的收藏狀態`);
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    
    if (!casesSheet) {
      throw new Error('找不到案件工作表');
    }
    
    // 尋找案件行
    const data = getSheetData(casesSheet);
    const rowIndex = data.findIndex(row => row.id === caseId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的案件');
    }
    
    // 獲取當前收藏狀態
    const currentRow = data[rowIndex - 2];
    Logger.log(`案件 ${caseId} 的原始收藏狀態為: ${currentRow.isFavorite}, 類型: ${typeof currentRow.isFavorite}`);
    const currentFavoriteStatus = String(currentRow.isFavorite).toLowerCase() === 'true';
    
    // 切換收藏狀態
    const newStatus = !currentFavoriteStatus;
    Logger.log(`案件 ${caseId} 的新收藏狀態為: ${newStatus}, 要寫入工作表的值: ${newStatus ? 'true' : 'false'}`);
    casesSheet.getRange(rowIndex, 9).setValue(newStatus ? 'true' : 'false');
    
    Logger.log(`案件 ${caseId} 的收藏狀態從 ${currentFavoriteStatus} 切換為 ${newStatus}`);
    
    return JSON.stringify({ 
      success: true, 
      isFavorite: newStatus,
      message: newStatus ? '已將案件加入收藏' : '已將案件移除收藏'
    });
  } catch (error) {
    console.error('切換案件收藏狀態時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 根據類別篩選案件
function filterCasesByCategory(sheetsId, categoryId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    // 如果沒有提供類別ID，則返回所有案件
    if (!categoryId) {
      return getAllCases(sheetsId, '');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    const progressSheet = ss.getSheetByName('進度');
    const relationsSheet = ss.getSheetByName('相關人關聯');
    const peopleSheet = ss.getSheetByName('相關人');
    const categoriesSheet = ss.getSheetByName('類別');
    const categoryRelationsSheet = ss.getSheetByName('類別關聯');
    
    if (!casesSheet || !categoryRelationsSheet) {
      throw new Error('找不到必要的工作表');
    }
    
    // 讀取案件資料
    const casesData = getSheetData(casesSheet);
    const progressData = getSheetData(progressSheet);
    const relationsData = getSheetData(relationsSheet);
    const peopleData = getSheetData(peopleSheet);
    const categoriesData = getSheetData(categoriesSheet);
    const categoryRelationsData = getSheetData(categoryRelationsSheet);
    
    // 根據類別關聯篩選案件
    const filteredCaseIds = categoryRelationsData
      .filter(relation => 
        relation.categoryId === categoryId && 
        relation.objectType === '案件')
      .map(relation => relation.objectId);
    
    Logger.log(`根據類別 ID ${categoryId} 篩選出 ${filteredCaseIds.length} 個相關案件`);
    
    // 篩選案件資料
    const filteredCases = [];
    const filteredCompletedCases = [];
    
    // 格式化相關人資料
    const people = peopleData.map(person => {
      return {
        id: person.id,
        name: person.name,
        color: person.color || generateRandomColor(),
        createDate: person.createDate || new Date().toISOString(),
        isActive: !(person.isActive === 'false' || person.isActive === false)
      };
    });
    
    // 格式化類別資料
    const categories = categoriesData.map(category => {
      return {
        id: category.id,
        name: category.name,
        color: category.color || generateRandomColor(),
        createDate: category.createDate || new Date().toISOString(),
        isActive: !(category.isActive === 'false' || category.isActive === false)
      };
    });
    
    // 處理案件資料
    for (const caseRow of casesData) {
      if (!filteredCaseIds.includes(caseRow.id)) {
        continue; // 跳過不相關的案件
      }
      
      try {
        // 從案件的 H 欄位讀取類別資訊
        let caseCategories = [];
        if (caseRow.categories && typeof caseRow.categories === 'string' && caseRow.categories.trim() !== '') {
          try {
            caseCategories = JSON.parse(caseRow.categories);
            Logger.log(`[案件 ID: ${caseRow.id}] 解析 H 欄位類別資訊: ${JSON.stringify(caseCategories)}`);
          } catch (jsonError) {
            Logger.log(`[案件 ID: ${caseRow.id}] 解析 H 欄位類別資訊時發生錯誤: ${jsonError}`);
          }
        }
        
        // 從關聯表中獲取該案件的相關人ID
        const caseRelations = relationsData.filter(relation => 
          relation.objectId === caseRow.id && relation.objectType === '案件');
        const relatedPeopleFromRelations = caseRelations.map(relation => relation.personId);
        
        // 檢查案件資料中是否有直接存儲的相關人資訊
        let relatedPeople = relatedPeopleFromRelations;
        if (caseRow.relatedPeople && typeof caseRow.relatedPeople === 'string' && caseRow.relatedPeople.trim() !== '') {
          try {
            const parsedRelatedPeople = JSON.parse(caseRow.relatedPeople);
            if (Array.isArray(parsedRelatedPeople) && parsedRelatedPeople.length > 0) {
              relatedPeople = parsedRelatedPeople; // 優先使用 G 欄位的相關人資訊
            }
          } catch (jsonError) {
            // 發生錯誤時使用關聯表中的相關人資訊
            Logger.log(`[案件 ID: ${caseRow.id}] 解析 G 欄位相關人資訊時發生錯誤: ${jsonError}`);
          }
        }
        
        // 解析檔案資訊
        let filesArray = [];
        if (caseRow.files && typeof caseRow.files === 'string') {
          try {
            filesArray = JSON.parse(caseRow.files);
          } catch (jsonError) {
            Logger.log(`[案件 ID: ${caseRow.id}] 解析檔案資訊時發生錯誤: ${jsonError}`);
            filesArray = [];
          }
        } else if (Array.isArray(caseRow.files)) {
          filesArray = caseRow.files;
        }
        
        const caseItem = {
          id: caseRow.id,
          title: caseRow.title,
          content: caseRow.content || '',
          date: caseRow.date || new Date().toISOString(),
          relatedPeople: relatedPeople,
          categories: caseCategories, // 使用解析後的類別資訊
          files: filesArray,
          progress: [],
          completed: caseRow.completed === 'true',
          isFavorite: caseRow.isFavorite === 'true'
        };
        
        // 添加進度資料
        if (progressData && progressData.length > 0) {
          for (const progressRow of progressData) {
            if (progressRow.caseId === caseItem.id) {
              // 從關聯表中獲取該進度的相關人ID
              const progressRelations = relationsData.filter(relation => 
                relation.objectId === progressRow.id && relation.objectType === '進度');
              const progressRelatedPeople = progressRelations.map(relation => relation.personId);
              
              let progressFilesArray = [];
              if (progressRow.files) {
                try {
                  progressFilesArray = JSON.parse(progressRow.files);
                } catch (jsonError) {
                  progressFilesArray = [];
                }
              }
              
              caseItem.progress.push({
                id: progressRow.id,
                date: progressRow.date || new Date().toISOString(),
                relatedPeople: progressRelatedPeople,
                content: progressRow.content || '',
                files: progressFilesArray,
                isDeleted: progressRow.isDeleted // 確保包含isDeleted屬性
              });
            }
          }
        }
        
        // 分類為進行中或已完成案件
        if (caseItem.completed) {
          filteredCompletedCases.push(caseItem);
        } else {
          filteredCases.push(caseItem);
        }
      } catch (caseError) {
        Logger.log(`處理案件 ${caseRow.id} 時發生錯誤: ${caseError}`);
        continue;
      }
    }
    
    Logger.log(`根據類別 ID ${categoryId} 篩選出 ${filteredCases.length} 個進行中案件和 ${filteredCompletedCases.length} 個已完成案件`);
    
    // 返回結果
    return JSON.stringify({
      success: true,
      cases: filteredCases,
      completedCases: filteredCompletedCases,
      people: people,
      categories: categories
    });
  } catch (error) {
    Logger.log(`按類別篩選案件時發生錯誤: ${error}`);
    return JSON.stringify({
      success: false,
      error: error.toString()
    });
  }
}

// 更新類別表結構，添加排序欄位
function updateCategoryTableStructure(sheetsId) {
  try {
    const ss = SpreadsheetApp.openById(sheetsId);
    const categoriesSheet = ss.getSheetByName('類別');
    
    if (!categoriesSheet) {
      return { success: false, error: '找不到類別工作表' };
    }
    
    // 獲取標題行
    const headers = categoriesSheet.getRange(1, 1, 1, categoriesSheet.getLastColumn()).getValues()[0];
    
    // 檢查是否已存在order欄位
    if (!headers.includes('order')) {
      // 添加新的order欄位
      categoriesSheet.getRange(1, headers.length + 1).setValue('order');
      
      // 獲取所有數據
      const data = getSheetData(categoriesSheet);
      
      // 為現有類別設置初始排序順序
      for (let i = 0; i < data.length; i++) {
        categoriesSheet.getRange(i + 2, headers.length + 1).setValue(i + 1);
      }
      
      return { success: true, message: '成功添加類別排序欄位' };
    }
    
    return { success: true, message: '類別排序欄位已存在' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// 批量更新類別順序
function updateCategoriesOrder(sheetsId, categoriesOrder) {
  try {
    if (!sheetsId || !categoriesOrder || !Array.isArray(categoriesOrder)) {
      throw new Error('未提供有效的排序資料');
    }
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const categoriesSheet = ss.getSheetByName('類別');
    
    if (!categoriesSheet) {
      throw new Error('找不到類別工作表');
    }
    
    // 確保存在order欄位
    const result = updateCategoryTableStructure(sheetsId);
    if (!result.success) {
      throw new Error('無法添加排序欄位: ' + result.error);
    }
    
    // 獲取標題行
    const headers = categoriesSheet.getRange(1, 1, 1, categoriesSheet.getLastColumn()).getValues()[0];
    const orderColIndex = headers.indexOf('order') + 1;
    
    if (orderColIndex < 1) {
      throw new Error('無法找到排序欄位');
    }
    
    // 讀取所有類別資料
    const data = getSheetData(categoriesSheet);
    
    // 更新每個類別的順序
    for (let i = 0; i < categoriesOrder.length; i++) {
      const categoryId = categoriesOrder[i];
      const rowIndex = data.findIndex(row => row.id === categoryId) + 2;
      
      if (rowIndex >= 2) {
        // 確保寫入的是數值而非字串
        categoriesSheet.getRange(rowIndex, orderColIndex).setValue(Number(i + 1));
        Logger.log(`更新類別 ID ${categoryId} 的順序為 ${i + 1}`);
      }
    }
    
    return JSON.stringify({ 
      success: true,
      message: '成功更新所有類別順序'
    });
  } catch (error) {
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 標記案件為已刪除
function markCaseAsDeleted(sheetsId, caseId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!caseId) {
      throw new Error('未提供案件 ID');
    }
    
    Logger.log(`標記案件 ${caseId} 為已刪除`);
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const casesSheet = ss.getSheetByName('案件');
    
    if (!casesSheet) {
      throw new Error('找不到案件工作表');
    }
    
    // 檢查是否存在isDeleted欄位，如果不存在則添加
    const headers = casesSheet.getRange(1, 1, 1, casesSheet.getLastColumn()).getValues()[0];
    let isDeletedColIndex = headers.indexOf('isDeleted') + 1;
    
    if (isDeletedColIndex < 1) {
      // 添加isDeleted欄位到標題行
      isDeletedColIndex = headers.length + 1;
      casesSheet.getRange(1, isDeletedColIndex).setValue('isDeleted');
      Logger.log(`在案件工作表中添加了isDeleted欄位，位置: ${isDeletedColIndex}`);
    }
    
    // 尋找案件行
    const data = getSheetData(casesSheet);
    const rowIndex = data.findIndex(row => row.id === caseId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的案件');
    }
    
    // 更新刪除狀態
    casesSheet.getRange(rowIndex, isDeletedColIndex).setValue(true);
    Logger.log(`已將案件 ${caseId} 標記為已刪除`);
    
    return JSON.stringify({ 
      success: true,
      message: '案件已標記為已刪除'
    });
  } catch (error) {
    console.error('標記案件為已刪除時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}

// 標記進度為已刪除
function markProgressAsDeleted(sheetsId, progressId) {
  try {
    if (!sheetsId) {
      throw new Error('未提供 Google Sheets ID');
    }
    
    if (!progressId) {
      throw new Error('未提供進度 ID');
    }
    
    Logger.log(`標記進度 ${progressId} 為已刪除`);
    
    const ss = SpreadsheetApp.openById(sheetsId);
    const progressSheet = ss.getSheetByName('進度');
    
    if (!progressSheet) {
      throw new Error('找不到進度工作表');
    }
    
    // 檢查是否存在isDeleted欄位，如果不存在則添加
    const headers = progressSheet.getRange(1, 1, 1, progressSheet.getLastColumn()).getValues()[0];
    let isDeletedColIndex = headers.indexOf('isDeleted') + 1;
    
    if (isDeletedColIndex < 1) {
      // 添加isDeleted欄位到標題行
      isDeletedColIndex = headers.length + 1;
      progressSheet.getRange(1, isDeletedColIndex).setValue('isDeleted');
      Logger.log(`在進度工作表中添加了isDeleted欄位，位置: ${isDeletedColIndex}`);
    }
    
    // 尋找進度行
    const data = getSheetData(progressSheet);
    const rowIndex = data.findIndex(row => row.id === progressId) + 2; // +2 是因為標題行和 0-index
    
    if (rowIndex < 2) {
      throw new Error('找不到指定的進度');
    }
    
    // 更新刪除狀態
    progressSheet.getRange(rowIndex, isDeletedColIndex).setValue(true);
    Logger.log(`已將進度 ${progressId} 標記為已刪除`);
    
    return JSON.stringify({ 
      success: true,
      message: '進度已標記為已刪除'
    });
  } catch (error) {
    console.error('標記進度為已刪除時發生錯誤:', error);
    return JSON.stringify({ 
      success: false, 
      error: error.toString() 
    });
  }
}
