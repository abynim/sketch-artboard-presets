function loadPresets(context) {
  try {
    // Enter insert artboard mode
    // (required for preset store to be initialized)
    const action = context.document
      .actionsController()
      .actionForID("MSInsertArtboardAction");
    action.insertArtboard(nil);

    // Load presets from store
    const store = MSArtboardPresetStore.alloc().init();
    store.loadUserPresets();
    const presetSections = store.valueForKeyPath(
      "categories.@unionOfArrays.sections"
    );

    // Recreate Commands and Menu for manifest
    let commands = [
      {
        script: "script.js",
        handler: "loadPresets",
        name: "Reload Presets",
        identifier: "loadPresets",
      },
    ];
    let menu = {
      title: "Artboard Presets",
      items: [],
    };

    // Populate commands and menu
    let loop = presetSections.objectEnumerator();
    let section;
    let count = 0;
    while ((section = loop.nextObject())) {
      let presets = section.presets();
      if (presets.count() > 0) {
        let preset;
        let loop2 = presets.objectEnumerator();

        const sectionName = section.name() || "Uncategorized";
        let sectionItem = {
          title: sectionName,
          items: [],
        };

        while ((preset = loop2.nextObject())) {
          const presetID = "preset" + ++count;
          sectionItem.items.push(presetID);

          commands.push({
            script: "script.js",
            handler: "insertArtboardFromPreset",
            name: preset.name(),
            identifier: presetID,
            dictionaryRepresentation: preset.dictionaryRepresentation(),
          });
        }
        menu.items.push(sectionItem);
      }
    }

    menu.items.push("-");
    menu.items.push("loadPresets");

    // Update manifest
    let manifestJSON = getManifestJSON(context);
    manifestJSON.commands = commands;
    manifestJSON.menu = menu;

    // Exit insert artboard mode
    action.insertArtboard(nil);

    // Save manifest
    const manifestURL = getManifestURL(context);
    NSString.alloc()
      .initWithData_encoding(
        NSJSONSerialization.dataWithJSONObject_options_error(
          manifestJSON,
          NSJSONWritingPrettyPrinted,
          nil
        ),
        NSUTF8StringEncoding
      )
      .writeToURL_atomically_encoding_error(
        manifestURL,
        true,
        NSUTF8StringEncoding,
        nil
      );

    // Reload plugins
    AppController.sharedInstance().pluginManager().reloadPlugins();
  } catch (error) {
    console.log("[Presets] Error loading presets: " + error);
  }
}

function insertArtboardFromPreset(context) {
  try {
    // Fetch command JSON
    const identifier = context.command.identifier();
    const commands = getManifestJSON(context).commands;
    const pred = NSPredicate.predicateWithFormat(
      "identifier == %@",
      identifier
    );
    const command = commands.filteredArrayUsingPredicate(pred).firstObject();

    // Enter insert artboard mode
    // (required for activating appropriate event handler)
    const action = context.document
      .actionsController()
      .actionForID("MSInsertArtboardAction");
    action.insertArtboard(nil);
    const evtHandler = context.document
      .inspectorController()
      .currentController()
      .eventHandler();

    // Make preset
    const preset = MSArtboardPreset.alloc().initWithDictionaryRepresentation(
      command.dictionaryRepresentation
    );

    // Insert artboard
    evtHandler.insertArtboardFromPreset(preset);

    // Exit insert artboard mode
    action.insertArtboard(nil);
  } catch (error) {
    console.log("[Presets] Error inserting artboard from preset: " + error);
  }
}

function getManifestURL(context) {
  return context.plugin
    .url()
    .URLByAppendingPathComponent("Contents")
    .URLByAppendingPathComponent("Sketch")
    .URLByAppendingPathComponent("manifest")
    .URLByAppendingPathExtension("json");
}

function getManifestJSON(context) {
  return NSJSONSerialization.JSONObjectWithData_options_error(
    NSData.dataWithContentsOfURL(getManifestURL(context)),
    NSJSONReadingMutableContainers,
    nil
  );
}
