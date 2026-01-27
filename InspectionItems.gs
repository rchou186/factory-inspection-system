/**
 * 點檢項目資料模組
 * Inspection Items Data Module
 * 
 * 包含所有11個區域、105個點檢項目 + 6個溫濕度記錄點
 */

/**
 * 取得所有點檢項目
 */
function getAllInspectionItems() {
  return [
    // 區域 1: 2F倉儲區 (3項)
    {
      code: 'AREA_2F',
      name_zh: '2F倉儲區',
      name_en: '2F Storage Area',
      items: [
        {no: 1, name_zh: '包材物品是否擺放整齊並關燈', name_en: 'Are materials placed neatly and lights off?'},
        {no: 2, name_zh: '充電式升高機，是否有歸位並充電', name_en: 'Is the rechargeable stacker returned and charging properly?'},
        {no: 3, name_zh: '地板是否有保持乾淨', name_en: 'Is the floor kept clean and dry?'}
      ]
    },
    
    // 區域 2: 研發實驗室 (2項)
    {
      code: 'AREA_RD',
      name_zh: '研發實驗室',
      name_en: 'R&D Laboratory',
      items: [
        {no: 1, name_zh: '研發室，是否有保持整齊清潔', name_en: 'Is the R&D room kept clean?'},
        {no: 2, name_zh: '實驗室，是否有保持整齊乾淨', name_en: 'Is the lab room kept clean?'}
      ]
    },
    
    // 區域 3: B1區 (10項，含2項溫濕度)
    {
      code: 'AREA_B1',
      name_zh: 'B1區',
      name_en: 'B1 Area',
      items: [
        {no: 1, name_zh: '物料是否擺放整齊', name_en: 'Are materials placed neatly?'},
        {no: 2, name_zh: '地板是否有保持乾淨', name_en: 'Is the floor kept clean and dry?'},
        {no: 3, name_zh: '生產工具，是否歸至定位', name_en: 'Production tools returned to designated position'},
        {no: 4, name_zh: '工具間，是否整齊', name_en: 'Tool room tidiness'},
        {no: 5, name_zh: '清潔用品庫，是否有整齊並上鎖關燈關門', name_en: 'Storage room: cleaned, locked, and lights off'},
        {no: 6, name_zh: '儲藏室，是否有清潔並上鎖關燈', name_en: 'Storage room: cleaned, locked, and lights off'},
        {no: 7, name_zh: '秤料室，是否有清潔並上鎖關燈', name_en: 'Weighing room: cleaned, locked, and door closed'},
        {no: 8, name_zh: '儲水槽，周邊是否清潔無積水', name_en: 'Water tank: cleaned, no water accumulation on the floor'},
        {no: 9, name_zh: '秤料室濕度 (%RH)', name_en: 'Weighing Room Humidity (%RH)', type: 'humidity'},
        {no: 10, name_zh: '物料室濕度 (%RH)', name_en: 'Material Storage Room Humidity (%RH)', type: 'humidity'}
      ]
    },
    
    // 區域 4: 更衣消毒室 (7項)
    {
      code: 'AREA_CHANGING',
      name_zh: '更衣消毒室',
      name_en: 'Changing & Disinfection Room',
      items: [
        {no: 1, name_zh: '垃圾桶是否清潔，並放置垃圾袋', name_en: 'Is the trash bin clean and with trash bags placed?'},
        {no: 2, name_zh: '浴塵室是否正常運作', name_en: 'Is the shower/restroom functioning properly?'},
        {no: 3, name_zh: '消毒池是否加藥，有效氯達200ppm', name_en: 'Is the disinfecting tank dosed with chlorine (effective chlorine ≥ 200 ppm)?'},
        {no: 4, name_zh: '洗手台是否清潔', name_en: 'Is the handwashing sink clean?'},
        {no: 5, name_zh: '酒精及洗手乳機，是否有充足及正常運作', name_en: 'Are the alcohol and soap dispensers filled and functioning properly?'},
        {no: 6, name_zh: '烘手機是否正常運作', name_en: 'Is the hand dryer functioning properly?'},
        {no: 7, name_zh: '地面是否維持清潔', name_en: 'Is the floor maintained clean?'}
      ]
    },
    
    // 區域 5: 緩衝區 (8項)
    {
      code: 'AREA_BUFFER',
      name_zh: '緩衝區',
      name_en: 'Buffer Zone',
      items: [
        {no: 1, name_zh: '天花板是否清潔', name_en: 'Is the ceiling clean?'},
        {no: 2, name_zh: '排風扇是否清潔，依規定設定開啟/關閉', name_en: 'Is the exhaust fan clean and set according to regulations?'},
        {no: 3, name_zh: '工具間是否清潔並擺放整齊', name_en: 'Is the tool room clean and items arranged neatly?'},
        {no: 4, name_zh: '冷氣是否清潔及排水正常', name_en: 'Is the air conditioner clean and draining properly?'},
        {no: 5, name_zh: '地面是否清潔無積水', name_en: 'Is the floor clean and free of standing water?'},
        {no: 6, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Are the drains and covers clean and rodent mesh in place?'},
        {no: 7, name_zh: '水管，是否整齊並歸位', name_en: 'Pipes: properly arranged and placed'},
        {no: 8, name_zh: '洗腳池,是否每日更換並收工後清理乾淨?', name_en: 'Foot bath: changed daily and cleaned after work?'}
      ]
    },
    
    // 區域 6: 包裝區 (17項)
    {
      code: 'AREA_PACKING',
      name_zh: '包裝區',
      name_en: 'Packing Area',
      items: [
        {no: 1, name_zh: '震動漏斗是否清潔，電源是否關閉', name_en: 'Vibrating hopper: cleaned and power off'},
        {no: 2, name_zh: '出口輸送帶是否清潔，電源是否關閉', name_en: 'Exit conveyor belt: cleaned and power off'},
        {no: 3, name_zh: '爬坡輸送帶是否清潔，電源是否關閉', name_en: 'Inclined conveyor belt: cleaned and power off'},
        {no: 4, name_zh: '杯式輸送機是否清潔，電源是否關閉', name_en: 'Cup conveyor: cleaned and power off'},
        {no: 5, name_zh: '計量秤是否清潔，電源是否關閉', name_en: 'Weighing scale: cleaned and power off'},
        {no: 6, name_zh: '包裝機是否清潔，電源是否關閉', name_en: 'Packing machine: cleaned and power off'},
        {no: 7, name_zh: '金檢機是否清潔，電源是否關閉', name_en: 'Metal detector: cleaned and power off'},
        {no: 8, name_zh: '捲類成型機是否清潔，電源是否關閉並歸位', name_en: 'Moulding machine: cleaned, power off and returned'},
        {no: 9, name_zh: '打釘機是否清潔，電源是否關閉並歸位', name_en: 'Stapler: cleaned, power off and returned'},
        {no: 10, name_zh: '拉膜機是否清潔，電源是否關閉', name_en: 'Film wrapping machine: cleaned and power off'},
        {no: 11, name_zh: '垃圾桶是否清淨並擺放整齊', name_en: 'Trash bins: cleaned and arranged properly'},
        {no: 12, name_zh: '地面是否清潔無積水', name_en: 'Floor: clean with no standing water'},
        {no: 13, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Drain and drain cover: cleaned and rodent mesh in place'},
        {no: 14, name_zh: '水管，是否整齊並歸位', name_en: 'Pipes: properly arranged and placed'},
        {no: 15, name_zh: '電梯口酒精,是否充足及正常運作', name_en: 'Elevator alcohol dispenser: sufficient and functioning'},
        {no: 16, name_zh: '水管及氣槍，是否整齊並歸位', name_en: 'Water pipes and air guns: arranged and returned'},
        {no: 17, name_zh: '地板膠帶,是否有更換?', name_en: 'Floor tape: replaced?'}
      ]
    },
    
    // 區域 7: 打漿區 (14項)
    {
      code: 'AREA_MIXING',
      name_zh: '打漿區',
      name_en: 'Mixing Area',
      items: [
        {no: 1, name_zh: '攪拌機是否清潔，電源是否關閉', name_en: 'Mixer: cleaned and power off'},
        {no: 2, name_zh: '細切機是否清潔，電源是否關閉', name_en: 'Fine cutter: cleaned and power off'},
        {no: 3, name_zh: '絞肉機是否清潔，電源是否關閉', name_en: 'Meat grinder: cleaned and power off'},
        {no: 4, name_zh: '天花板、牆壁是否清潔', name_en: 'Ceiling and walls cleanliness'},
        {no: 5, name_zh: '半成品庫門、門簾、地面是否清潔', name_en: 'Semi-product storage doors, curtains, and floor cleanliness'},
        {no: 6, name_zh: '半成品庫內原物料是否正確擺放', name_en: 'Semi-product storage: raw materials placed correctly'},
        {no: 7, name_zh: '管線零件是否清潔(含籃子)', name_en: 'Pipeline components: cleaned (including mats)'},
        {no: 8, name_zh: '地秤是否清潔，電源是否關閉', name_en: 'Floor scale: cleaned and power off'},
        {no: 9, name_zh: '地面是否清潔無積水', name_en: 'No standing water on the floor'},
        {no: 10, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Drain and drain cover: cleaned and rodent mesh in place'},
        {no: 11, name_zh: '生產器具，是否清潔並擺放整齊', name_en: 'Production tools: cleaned and arranged properly'},
        {no: 12, name_zh: '漿桶，是否清潔並擺放整齊', name_en: 'Ingredient buckets: cleaned and arranged properly'},
        {no: 13, name_zh: '水管及氣槍是否整齊並歸位', name_en: 'Pipes and air guns: properly arranged and placed'},
        {no: 14, name_zh: '地板膠帶,是否有更換?', name_en: 'Floor tape: replaced?'}
      ]
    },
    
    // 區域 8: 前處理區 (9項)
    {
      code: 'AREA_PREP',
      name_zh: '前處理區',
      name_en: 'Preparation Area',
      items: [
        {no: 1, name_zh: '刨片機是否清潔，電源及空壓是否關閉', name_en: 'Is the slicer clean and power/air pressure turned off?'},
        {no: 2, name_zh: '垃圾桶是否清潔', name_en: 'Is the trash bin clean?'},
        {no: 3, name_zh: '地面是否清潔無積水', name_en: 'Is the floor clean and free of standing water?'},
        {no: 4, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Are the drains and covers clean and rodent mesh in place?'},
        {no: 5, name_zh: '冷氣及濾網是否清潔及排水正常', name_en: 'Is the air conditioner clean and draining properly?'},
        {no: 6, name_zh: '水管及氣槍是否整齊並歸位', name_en: 'Are pipes tidy and returned to place?'},
        {no: 7, name_zh: '零件是否清潔(含籃子)', name_en: 'Are parts clean (including mats)?'},
        {no: 8, name_zh: '排風扇是否清潔，依規定設定開啟/關閉', name_en: 'Is the exhaust fan clean and set according to regulations?'},
        {no: 9, name_zh: '地板膠帶,是否有更換?', name_en: 'Floor tape: replaced?'}
      ]
    },
    
    // 區域 9: 搥漿區 (7項)
    {
      code: 'AREA_PASTE',
      name_zh: '搥漿區',
      name_en: 'Paste Processing Area',
      items: [
        {no: 1, name_zh: '搥漿機是否清潔，電源是否關閉', name_en: 'Paste dispenser: cleaned and power off'},
        {no: 2, name_zh: '天花板及燈具是否清潔，電源是否關閉', name_en: 'Ceiling and lighting fixtures: cleaned and power off'},
        {no: 3, name_zh: '拉門及牆壁是否清潔', name_en: 'Sliding door cleanliness'},
        {no: 4, name_zh: '地面是否清潔並無積水', name_en: 'Floor: clean with no standing water'},
        {no: 5, name_zh: '排風扇是否清潔，電源是否關閉', name_en: 'Exhaust fan: cleaned and power off'},
        {no: 6, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Drain and drain cover: cleaned with rodent mesh in place'},
        {no: 7, name_zh: '生產工具,是否有擺放整齊?', name_en: 'Production tools: arranged properly?'}
      ]
    },
    
    // 區域 10: 蒸煮區 (15項，含2項溫度)
    {
      code: 'AREA_COOKING',
      name_zh: '蒸煮區',
      name_en: 'Cooking Area',
      items: [
        {no: 1, name_zh: '充填機是否清淨，電源是否關閉', name_en: 'Is the filling machine clean and power off?'},
        {no: 2, name_zh: '管道式金檢機是否清潔', name_en: 'Is the pipeline-type metal detector clean?'},
        {no: 3, name_zh: '成型機是否清潔，電源是否關閉', name_en: 'Is the forming machine clean and power off?'},
        {no: 4, name_zh: '蒸煮機、洗滌槽是否清潔，電源是否關閉', name_en: 'Steamer and washing tank: cleaned and power off'},
        {no: 5, name_zh: '零件是否清潔並正確擺放', name_en: 'Are parts clean and properly placed?'},
        {no: 6, name_zh: '蒸煮機兩側塑膠片是否乾淨', name_en: 'Are the plastic panels on both sides of the steamer clean?'},
        {no: 7, name_zh: '牆面及地面是否清潔無積水', name_en: 'Is the floor clean and free of standing water?'},
        {no: 8, name_zh: '預冷機總成是否清潔，電源是否關閉', name_en: 'Is the pre-cooling unit clean and power off?'},
        {no: 9, name_zh: '預冷機風扇是否依規定吹風', name_en: 'Is the pre-cooling fan operating as required?'},
        {no: 10, name_zh: '水溝及水溝蓋是否清潔並放置鼠網', name_en: 'Are the drains and covers clean and rodent mesh in place?'},
        {no: 11, name_zh: 'IQF是否清潔（含出入口及擋板）', name_en: 'IQF: cleaned (including entry, exit and baffles)'},
        {no: 12, name_zh: 'IQF除霜開關是否關閉', name_en: 'IQF defrost switch: turned off'},
        {no: 13, name_zh: '水管及氣槍，是否整齊並歸位', name_en: 'Pipes: properly arranged and placed'},
        {no: 14, name_zh: '半成品暫存區溫度 (℃)', name_en: 'Semi-product temporary storage temperature (℃)', type: 'temperature'},
        {no: 15, name_zh: '半成品冷凍區溫度 (℃)', name_en: 'Semi-product frozen storage temperature (℃)', type: 'temperature'}
      ]
    },
    
    // 區域 11: 其他 (13項，含2項溫度)
    {
      code: 'AREA_OTHER',
      name_zh: '其他',
      name_en: 'Others',
      items: [
        {no: 1, name_zh: '一樓廁所是否清潔', name_en: 'Is the 1F restroom clean?'},
        {no: 2, name_zh: '二樓廁所是否清潔', name_en: 'Is the 2F restroom clean?'},
        {no: 3, name_zh: '清潔用品庫是否整齊並上鎖', name_en: 'Is the cleaning supply room tidy and locked?'},
        {no: 4, name_zh: '廢水區是否正常運轉', name_en: 'Is the wastewater area operating normally?'},
        {no: 5, name_zh: '攔汙閘是否保持清潔(裡與外)', name_en: 'Is the changing area clean (inside and outside)?'},
        {no: 6, name_zh: '大門樓梯是否無積水並保持清潔', name_en: 'Are the front entrance stairs clean?'},
        {no: 7, name_zh: '後門樓梯及牆面是否清潔', name_en: 'Are the rear entrance stairs clean?'},
        {no: 8, name_zh: '垃圾車周圍環境是否清潔', name_en: 'Is the garbage area and its surroundings clean?'},
        {no: 9, name_zh: '電梯內部是否清潔，隨手關門並停放２樓', name_en: 'Is the elevator clean, door closed, and properly parked on 2F?'},
        {no: 10, name_zh: '冰庫溫度是否正常，壓縮機是否開啟', name_en: 'Is the freezer temperature normal and compressor running?'},
        {no: 11, name_zh: '廠內鹽碗/風鈴/八卦鐘，是否擺放正確', name_en: 'Are the internal items placed correctly?'},
        {no: 12, name_zh: '原料庫溫度 (℃)', name_en: 'Raw material storage temperature (℃)', type: 'temperature'},
        {no: 13, name_zh: '成品庫溫度 (℃)', name_en: 'Finished product storage temperature (℃)', type: 'temperature'}
      ]
    }
  ];
}

/**
 * 取得指定區域的項目
 */
function getAreaItems(areaCode) {
  const allItems = getAllInspectionItems();
  return allItems.find(area => area.code === areaCode);
}

/**
 * 取得所有溫濕度測點
 */
function getAllTempHumidityPoints() {
  const points = [];
  const allItems = getAllInspectionItems();
  
  allItems.forEach(area => {
    area.items.forEach(item => {
      if (item.type === 'temperature' || item.type === 'humidity') {
        points.push({
          area: area.name_zh,
          areaCode: area.code,
          itemNo: item.no,
          name_zh: item.name_zh,
          name_en: item.name_en,
          type: item.type
        });
      }
    });
  });
  
  return points;
}
