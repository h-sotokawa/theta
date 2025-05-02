// メイン実行関数
function runHimeji() {
  const config = Config.getLocationConfig('himeji');
  transferDataMain(config.sourceId, config.destinationId, config.location);
}
