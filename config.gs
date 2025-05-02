const Config = {
  // 拠点ごとの設定を管理
  LOCATIONS: {
    osaka: {
      sources: [
        'SPREADSHEET_ID_SOURCE_OSAKA_DESKTOP',
        'SPREADSHEET_ID_SOURCE_OSAKA_LAPTOP'
      ],
      destination: 'SPREADSHEET_ID_DESTINATION',
      location: '大阪',
      prefix: 'OSAKA'
    },
    kobe: {
      source: 'SPREADSHEET_ID_SOURCE_KOBE',
      destination: 'SPREADSHEET_ID_DESTINATION',
      location: '神戸',
      prefix: 'Kobe'
    },
    himeji: {
      source: 'SPREADSHEET_ID_SOURCE_HIMEJI',
      destination: 'SPREADSHEET_ID_DESTINATION',
      location: '姫路',
      prefix: 'Hime'
    }
  },

  // 拠点ごとの設定を取得
  getLocationConfig: function(locationKey) {
    const scriptProperties = PropertiesService.getScriptProperties();
    const locationConfig = this.LOCATIONS[locationKey];
    
    if (!locationConfig) {
      throw new Error(`指定された拠点「${locationKey}」の設定が見つかりません。`);
    }

    if (locationKey === 'osaka') {
      // 大阪の場合は複数のソースIDを取得
      const sourceIds = locationConfig.sources.map(source => scriptProperties.getProperty(source));
      return {
        sourceIds: sourceIds,
        destinationId: scriptProperties.getProperty(locationConfig.destination),
        location: locationConfig.location,
        prefix: locationConfig.prefix
      };
    } else {
      // その他の拠点は単一のソースID
      return {
        sourceId: scriptProperties.getProperty(locationConfig.source),
        destinationId: scriptProperties.getProperty(locationConfig.destination),
        location: locationConfig.location,
        prefix: locationConfig.prefix
      };
    }
  },

  // エラーメール通知先を取得
  getErrorNotificationEmail: function() {
    const scriptProperties = PropertiesService.getScriptProperties();
    return scriptProperties.getProperty('ERROR_NOTIFICATION_EMAIL');
  }
}; 