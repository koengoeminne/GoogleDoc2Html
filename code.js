function ConvertGoogleDocToCleanHtml() {
  var body = DocumentApp.getActiveDocument().getBody();
  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }

  var html = output.join('\r\n');
  emailHtml(html, images);
  sendLog();
}

function sendLog() {
  var fname = "log";
  var files = DriveApp.getFilesByName(fname);
  var doc = (files.hasNext()) ? DocumentApp.openById(files.next().getId()) : DocumentApp.create(fname);
  doc.getBody().editAsText().appendText(Logger.getLog());
}

function emailHtml(html, images) {
  var attachments = [];
  for (var j=0; j<images.length; j++) {
    attachments.push( {
      "fileName": images[j].name,
      "mimeType": images[j].type,
      "content": images[j].blob.getBytes() } );
  }

  var inlineImages = {};
  for (var j=0; j<images.length; j++) {
    inlineImages[[images[j].name]] = images[j].blob;
  }

  var name = DocumentApp.getActiveDocument().getName()+".html";
  attachments.push({"fileName":name, "mimeType": "text/html", "content": html});
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: html,
     inlineImages: inlineImages,
     attachments: attachments
   });
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";
  
  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
      case DocumentApp.ParagraphHeading.HEADING6: 
        prefix = "<h6>", suffix = "</h6>"; break;
      case DocumentApp.ParagraphHeading.HEADING5: 
        prefix = "<h5>", suffix = "</h5>"; break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>"; break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>"; break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>"; break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>"; break;
      default: 
        prefix = "<p>", suffix = "</p>";
    }
    if (item.getNumChildren() == 0)
      return "";
  }
  else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
    processImage(item, images, output);
  }
  else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
    var gt = item.getGlyphType();
    var key = item.getListId() + '.' + item.getNestingLevel();
    var counter = listCounters[key] || 0;

    // First list item
    if ( counter == 0 ) {
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix = '<ul><li>';
        suffix = "</li>";
      } else {
        // Ordered list (<ol>):
        prefix = "<ol><li>", suffix = "</li>";
      }
    }
    else {
      prefix = "<li>";
      suffix = "</li>";
    }

    if (item.isAtDocumentEnd() || (item.getNextSibling() && (item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM))) {
      if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        suffix += "</ul>";
      }
      else {
        // Ordered list (<ol>):
        suffix += "</ol>";
      }
    }

    counter++;
    listCounters[key] = counter;
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processTextItem(item, output);
  }
  else {
    if (item.getNumChildren()) {
      for (var i = 0; i < item.getNumChildren(); i++) {
        output.push(processItem(item.getChild(i), listCounters, images));
      }
    }
  }

  output.push(suffix);
  return output.join('');
}

function processTextItem(item, output) {
  sanitizeTextItem(item);
  
  var text = item.getText();
  var indices = item.getTextAttributeIndices();
  
  if(indices.length > 1) {
    for (var i=0; i<indices.length; i++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      var partText = text.substring(startPos, endPos);
      processText(partText, partAtts, output);
    }
  }
  else {
    processText(text, item.getAttributes(), output);
  }
}

function sanitizeTextItem(item) {
  // Sanitize text from flickr garbage
  item.replaceText('<script.*/script>', '');
  item.replaceText('data-flickr-embed="true" ', 'target="_blank"');
  item.replaceText('(<img)', '<img class="aligncenter"');  
}

function processText(text, attr, output) {
  var prefix = "", suffix = "";
  
  Logger.log(text);
  for(var att in attr) {
    Logger.log(att + " : " + attr[att]);
  }
  
  // Possible attribute keys:
  // FONT_SIZE : ITALIC : STRIKETHROUGH : FOREGROUND_COLOR : BOLD : LINK_URL : UNDERLINE : FONT_FAMILY : BACKGROUND_COLOR
  if(attr['BOLD']) {
    prefix += "<strong>", suffix = "</strong>" + suffix;
  }
  if(attr['ITALIC']) {
    prefix += "<i>", suffix = "</i>" + suffix;
  }
  if(attr['UNDERLINE'] && !attr['LINK_URL']) {
    prefix += "<u>", suffix = "</u>" + suffix;
  }
  if(attr['LINK_URL']) {
    prefix += "<a href=" + attr['LINK_URL'] + " target=\"_blank\">", suffix = "</a>" + suffix;
  }
  
  output.push(prefix + text + suffix);
}

function processImage(item, images, output) {
  images = images || [];
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: "+contentType;
  }
  var imagePrefix = "Image_";
  var imageCounter = images.length;
  var name = imagePrefix + imageCounter + extension;
  imageCounter++;
  output.push('<img src="cid:'+name+'" />');
  images.push( {
    "blob": blob,
    "type": contentType,
    "name": name});
}