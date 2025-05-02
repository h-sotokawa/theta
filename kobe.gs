// メイン実行関数
function runKobe() {
  const config = Config.getLocationConfig('kobe');
  transferDataMain(config.sourceId, config.destinationId, config.location);
}

