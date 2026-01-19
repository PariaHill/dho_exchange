/**
 * 대항해시대 오리진 - 교환 트리 계산기
 * Google Apps Script 메인 코드
 */

// 웹앱 진입점
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('교환 트리 계산기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Exchanges 시트 데이터 가져오기
 */
function getExchangesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exchangesSheet = ss.getSheetByName('Exchanges');
  if (!exchangesSheet) return {};
  
  const data = exchangesSheet.getDataRange().getValues();
  const exchanges = {};
  
  // 헤더 제외
  for (let i = 1; i < data.length; i++) {
    const itemName = data[i][0]; // 물물교환품목
    const materialsStr = data[i][1]; // 재료 (/ 구분)
    const quantitiesStr = data[i][2]; // 수량 (/ 구분)
    
    if (!itemName || !materialsStr || !quantitiesStr) continue;
    
    // 재료 파싱
    const materials = String(materialsStr).split('/').map(m => m.trim());
    
    // 수량 파싱: 첫번째=결과물, 나머지=재료별 필요량
    const quantities = String(quantitiesStr).split('/').map(q => parseFloat(q.trim()));
    
    const outputQuantity = quantities[0] || 1; // 결과물 수량
    const materialQuantities = quantities.slice(1); // 재료별 필요 수량
    
    exchanges[itemName] = {
      name: itemName,
      outputQuantity: outputQuantity,
      materials: materials.map((mat, idx) => ({
        name: mat,
        quantity: materialQuantities[idx] || 1
      }))
    };
  }
  
  return exchanges;
}

/**
 * 완성품 목록 가져오기 (Exchanges에서)
 */
function getFinishedItems() {
  const exchanges = getExchangesData();
  return Object.keys(exchanges).map(name => ({
    name: name,
    type: '완성품'
  }));
}

/**
 * 특정 아이템 제작에 필요한 모든 재료 계산 (새 구조)
 */
function calculateMaterials(targetItemName, exchangeCount) {
  const exchanges = getExchangesData();
  const trades = getTradeData();
  const towns = getTownData();
  
  // 최종 재료 집계 (기본 재료 + 중간 재료)
  const totalMaterials = {};
  
  /**
   * 재귀적으로 재료 계산
   */
  function traverse(itemName, exchangeCount, depth, path) {
    // 순환 참조 방지
    if (path.includes(itemName)) {
      return {
        name: itemName,
        exchangeCount: exchangeCount,
        outputQuantity: 0,
        isCircular: true,
        depth: depth,
        children: []
      };
    }
    
    const newPath = [...path, itemName];
    const exchange = exchanges[itemName];
    
    // Exchanges에 레시피가 없으면 기본 재료
    if (!exchange || !exchange.materials || exchange.materials.length === 0) {
      if (!totalMaterials[itemName]) {
        // Trade에서 도시 정보 찾기
        const tradeInfo = trades[itemName] || {};
        const cities = tradeInfo.cities || [];
        
        totalMaterials[itemName] = {
          name: itemName,
          quantity: 0,
          location: cities.join(', '),
          isBase: true,
          towns: towns[itemName] || []
        };
      }
      // 기본 재료는 필요 수량 누적
      totalMaterials[itemName].quantity += exchangeCount;
      
      return {
        name: itemName,
        exchangeCount: exchangeCount,
        outputQuantity: exchangeCount,
        location: totalMaterials[itemName].location,
        isBase: true,
        depth: depth,
        children: []
      };
    }
    
    // 하위 재료가 있는 경우 (중간 재료)
    const outputQuantity = exchange.outputQuantity || 1;
    const totalOutput = exchangeCount * outputQuantity; // 총 생산량
    
    // 중간 재료도 집계에 추가 (depth > 0일 때만, 즉 루트 제외)
    if (depth > 0) {
      if (!totalMaterials[itemName]) {
        totalMaterials[itemName] = {
          name: itemName,
          quantity: 0,
          exchangeCount: 0,
          location: '',
          isBase: false,
          outputQuantity: outputQuantity,
          towns: towns[itemName] || []
        };
      }
      totalMaterials[itemName].quantity += totalOutput; // 총 생산량
      totalMaterials[itemName].exchangeCount += exchangeCount; // 교환 횟수
    }
    
    const children = [];
    exchange.materials.forEach(mat => {
      // 필요한 재료의 총량 계산
      const neededMaterialQuantity = mat.quantity * exchangeCount;
      
      // 하위 재료의 레시피 확인 (Exchanges에서)
      const subExchange = exchanges[mat.name];
      if (subExchange) {
        // 하위 재료도 교환으로 만들어야 하는 경우
        const subOutputQuantity = subExchange.outputQuantity || 1;
        // 필요한 교환 횟수 계산 (올림)
        const neededExchanges = Math.ceil(neededMaterialQuantity / subOutputQuantity);
        const childNode = traverse(mat.name, neededExchanges, depth + 1, newPath);
        children.push(childNode);
      } else {
        // 기본 재료인 경우 - 필요 수량이 곧 필요량
        const childNode = traverse(mat.name, neededMaterialQuantity, depth + 1, newPath);
        children.push(childNode);
      }
    });
    
    return {
      name: itemName,
      exchangeCount: exchangeCount,
      outputQuantity: totalOutput,
      location: '',
      isBase: false,
      depth: depth,
      children: children
    };
  }
  
  // 최대 깊이 제한 (무한 재귀 방지)
  const maxDepth = 6;
  const rootNode = traverse(targetItemName, exchangeCount, 0, []);
  
  return {
    tree: rootNode,
    totalMaterials: Object.values(totalMaterials)
  };
}

/**
 * Town 시트 데이터 가져오기
 */
function getTownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const townSheet = ss.getSheetByName('Town');
  if (!townSheet) return {};
  
  const data = townSheet.getDataRange().getValues();
  const towns = {};
  
  // 헤더 제외
  for (let i = 1; i < data.length; i++) {
    const townName = data[i][3]; // 이름 컬럼 (D열, 인덱스 3)
    const items = data[i][2]; // 물물교환품목 (C열, 인덱스 2)
    
    if (!townName || !items) continue;
    
    // 품목별로 역인덱스 생성
    const itemList = String(items).split('/').map(item => item.trim());
    
    itemList.forEach(itemName => {
      if (!towns[itemName]) {
        towns[itemName] = [];
      }
      // 중복 체크: 같은 마을이 이미 있으면 추가하지 않음
      if (!towns[itemName].includes(townName)) {
        towns[itemName].push(townName);
      }
    });
  }
  
  return towns;
}

/**
 * Port 시트 데이터 가져오기
 */
function getPortData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portSheet = ss.getSheetByName('Port');
  if (!portSheet) return {};
  
  const data = portSheet.getDataRange().getValues();
  const ports = {};
  
  // 헤더 제외
  for (let i = 1; i < data.length; i++) {
    const cityName = data[i][4]; // 이름 컬럼 (E열, 인덱스 4)
    if (!cityName) continue;
    
    ports[cityName] = {
      name: cityName,
      country: data[i][5] || '', // 국가
      region: data[i][6] || '', // 권역
      epidemic: data[i][7] || '', // 대유행
      climate: data[i][9] || '' // 기후
    };
  }
  
  return ports;
}

/**
 * Trade 시트 데이터 가져오기
 */
function getTradeData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradeSheet = ss.getSheetByName('Trade');
  if (!tradeSheet) return {};
  
  const data = tradeSheet.getDataRange().getValues();
  const trades = {};
  
  // 헤더 제외
  for (let i = 1; i < data.length; i++) {
    const itemName = data[i][0]; // 이름 (품목명)
    if (!itemName) continue;
    
    const cities = data[i][3] ? String(data[i][3]).split('/') : []; // 도시 (/ 구분)
    
    trades[itemName] = {
      name: itemName,
      category: data[i][1] || '', // 품목
      specialty: data[i][2] || '', // 명산품
      cities: cities.map(city => city.trim()), // 공백 제거
      peakSeason: data[i][4] || '', // 성수기
      normalSeason: data[i][5] || '', // 평수기
      offSeason: data[i][6] || '' // 비수기
    };
  }
  
  return trades;
}

/**
 * Season 시트 데이터 가져오기
 */
function getSeasonData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const seasonSheet = ss.getSheetByName('Season');
  if (!seasonSheet) return {};
  
  const data = seasonSheet.getDataRange().getValues();
  const seasons = {};
  
  // 헤더 제외
  for (let i = 1; i < data.length; i++) {
    const category = data[i][0]; // 품목
    const climate = data[i][1]; // 기후
    
    if (!category || !climate) continue;
    
    const key = category + '|' + climate;
    
    // 1월~12월 데이터 (컬럼 2~13)
    seasons[key] = {
      category: category,
      climate: climate,
      months: []
    };
    
    for (let month = 0; month < 12; month++) {
      seasons[key].months.push(data[i][2 + month] || '');
    }
  }
  
  return seasons;
}

/**
 * 현재 게임 내 월 계산
 */
function getCurrentGameMonth() {
  const now = new Date();
  // 한국 시간으로 변환
  const koreaTime = new Date(now.toLocaleString("en-US", {timeZone: "Asia/Seoul"}));
  
  // 오전 9시 기준으로 판단
  const hour = koreaTime.getHours();
  if (hour < 9) {
    // 아직 9시 전이면 어제 날짜 기준
    koreaTime.setDate(koreaTime.getDate() - 1);
  }
  
  // 기준일 설정 (2026.1.16 = 10월)
  const baseDate = new Date(2026, 0, 16, 9, 0, 0); // 2026년 1월 16일 9시
  const baseMonth = 10; // 10월
  
  // 경과 일수 계산
  const daysDiff = Math.floor((koreaTime - baseDate) / (1000 * 60 * 60 * 24));
  
  // 게임 내 월 계산 (12개월 순환)
  let gameMonth = ((baseMonth - 1 + daysDiff) % 12) + 1;
  if (gameMonth <= 0) gameMonth += 12;
  
  return gameMonth;
}

/**
 * 도시 이름 매칭 (띄어쓰기 무시)
 */
function matchCityName(cityName, portName) {
  if (!cityName || !portName) return false;
  
  // 1. 정확히 일치
  if (cityName === portName) return true;
  
  // 2. 공백 제거 후 비교
  const cleanCity = cityName.replace(/\s/g, '');
  const cleanPort = portName.replace(/\s/g, '');
  if (cleanCity === cleanPort) return true;
  
  // 3. 부분 문자열 매칭
  if (portName.includes(cityName) || cityName.includes(portName)) return true;
  
  return false;
}

/**
 * 재료의 성수기/비수기 정보 가져오기
 */
function getSeasonInfo(materialName) {
  const ports = getPortData();
  const trades = getTradeData();
  const seasons = getSeasonData();
  const currentMonth = getCurrentGameMonth();
  
  // Trade 시트에서 재료 정보 찾기
  const tradeInfo = trades[materialName];
  if (!tradeInfo) {
    return null;
  }
  
  const result = {
    currentMonth: currentMonth,
    cities: []
  };
  
  // 각 도시별로 시즌 정보 계산
  tradeInfo.cities.forEach(cityName => {
    // Port 시트에서 도시 찾기 (띄어쓰기 무시)
    let portInfo = null;
    for (let portName in ports) {
      if (matchCityName(cityName, portName)) {
        portInfo = ports[portName];
        break;
      }
    }
    
    if (!portInfo) {
      result.cities.push({
        name: cityName,
        season: null,
        climate: '알 수 없음',
        country: '',
        region: '',
        epidemic: ''
      });
      return;
    }
    
    // Season 시트에서 시즌 정보 찾기
    const seasonKey = tradeInfo.category + '|' + portInfo.climate;
    const seasonInfo = seasons[seasonKey];
    
    let seasonType = '평수기'; // 기본값
    
    if (seasonInfo && seasonInfo.months[currentMonth - 1]) {
      const value = seasonInfo.months[currentMonth - 1];
      if (value === '성수기') seasonType = '성수기';
      else if (value === '비수기') seasonType = '비수기';
      else seasonType = '평수기';
    }
    
    result.cities.push({
      name: cityName,
      season: seasonType,
      climate: portInfo.climate,
      country: portInfo.country,
      region: portInfo.region,
      epidemic: portInfo.epidemic
    });
  });
  
  return result;
}

/**
 * 웹앱에서 호출할 함수들
 */
function getItemList() {
  return getFinishedItems();
}

function calculate(itemName, quantity) {
  return calculateMaterials(itemName, quantity);
}

function getMaterialSeasonInfo(materialName) {
  return getSeasonInfo(materialName);
}
