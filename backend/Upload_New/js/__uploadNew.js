var uploadNew = {

  option : {
    rootPath : 'Upload_New/',
    previewDocIcon : 'Upload_New/img/doc/',
    noImage : 'Upload_New/img/no-image.gif',
    noFile : 'Upload_New/img/doc/no-file.svg',
    uploadServerSide : 'Upload_New/uploadNew.php'
  },

  init: function(){

    uploadNew.events();
    uploadNew.modal.events();

  },

  events: function(){

    let inputsFiles = document.getElementsByClassName('upload--file'),
        deleteFiles = document.getElementsByClassName('upload--delete-file');

    //
    Array.prototype.forEach.call(inputsFiles, function(e, i){

      let wrapperFile = e.parentNode.parentNode;

      //console.log(wrapperFile);
      e.addEventListener('change',function(){

        //If is an image file -The upload begin after validate dimensions in the "validateImageDimension" method
        //Note: validateImageDimension it's an asynchronous method
        if(uploadNew.fileIsImage(e.value)){

          if(uploadNew.validateExtension(e) && uploadNew.validateSize(e)) uploadNew.validateImageDimension(e, wrapperFile);
        
        //If it isn't an image -The upload begin after validations return "true"
        }else{
          
          if(uploadNew.validateExtension(e) && uploadNew.validateSize(e)) uploadNew.upload(e, wrapperFile);

        }
      
      });

      //Set ui on start
      uploadNew.ui.set(e, i);


    });
    //

    //
    Array.prototype.forEach.call(deleteFiles, function(e, i){

      e.addEventListener('click',function(){

          uploadNew.deleteFile(e, true);
      
      });

    });
    //

    //
    uploadNew.msjScreen.events();

  },

  //Validate the extension seted in data info -"data-file-accept" attribute in the input file
  validateExtension : function(e){

    let suportExtensions = e.getAttribute('data-file-accept') ? String(e.getAttribute('data-file-accept')).split(',') : false;

    uploadNew.msjScreen.delete(e);

    if(e.type == "file" && suportExtensions) {
        var sFileName = e.value;
        if (sFileName.length > 0) {
          var blnValid = false;
          for (var j = 0; j < suportExtensions.length; j++) {
              var sCurExtension = suportExtensions[j];
              if (sFileName.substr(sFileName.length - sCurExtension.length, sCurExtension.length).toLowerCase() == sCurExtension.toLowerCase()) {
                  blnValid = true;
                  break;
              }
          }
            
          if (!blnValid) {
              let msj = "<span>ATENCIÓN:</span>Disculpa, <strong>" + sFileName + "</strong> no es un tipo válido, solo se permiten archivos de tipo: <strong>" + suportExtensions.join(", ") + "</strong>";
              uploadNew.msjScreen.show(e, msj, 'error');
              e.value = "";
              return false;
          }
        }
    }
    return true;

  },

  //Validate the size seted in data info -"data-file-maxsize" attribute in the input file
  validateSize : function(e){

    if(!e.files[0]) return false;
    
    let suportSize = (e.getAttribute('data-file-maxsize') && e.getAttribute('data-file-maxsize') != '') ? e.getAttribute('data-file-maxsize') : false,
        fileSize = e.files[0].size / 1024 / 1024;

    uploadNew.msjScreen.delete(e);
    
    fileSize = fileSize.toFixed(2);

    if(fileSize && suportSize && (Number(fileSize) > Number(suportSize))){

      let sFileName = e.value,
          msj = "<span>ATENCIÓN:</span>Disculpa, " + sFileName + " de <strong>" +fileSize+ "Mb</strong> exede el peso máximo, solo son validos alchivos menores a: <strong>" +  suportSize + "Mb</strong>";

      uploadNew.msjScreen.show(e, msj, 'error');
      return false;

    }else{

      return true;

    }

  },

  //Validate the size (width and height) seted in data info -"data-file-width or/and data-file-height" attribute in the input file
  //This method is asynchronous, response afeter img.onload 
  validateImageDimension : function(e, wrapperFile){

    let _URL = window.URL || window.webkitURL, 
        file, 
        img,
        suportWidth = e.getAttribute('data-file-width') ? e.getAttribute('data-file-width') : false,
        suportHeight = e.getAttribute('data-file-height') ? e.getAttribute('data-file-height') : false,
        suportMaxWidth = e.getAttribute('data-file-maxwidth') ? e.getAttribute('data-file-maxwidth') : false,
        suportMaxHeight = e.getAttribute('data-file-maxheight') ? e.getAttribute('data-file-maxheight') : false,
        suportMinWidth = e.getAttribute('data-file-minwidth') ? e.getAttribute('data-file-minwidth') : false,
        suportMinHeight = e.getAttribute('data-file-minheight') ? e.getAttribute('data-file-minheight') : false;

    uploadNew.msjScreen.delete(e);

    if((file = e.files[0])){

        let img = new Image(),
            objectUrl = _URL.createObjectURL(file);

        img.onload = function(){
 
          let msj = '';

          if(suportWidth && this.width != suportWidth){ msj += "El ancho debe ser de: <strong>" + suportWidth + "px</strong><br>"; } 
          if(suportHeight && this.height != suportHeight){ msj += "El alto debe ser de: <strong>" + suportHeight + "px</strong><br>"; }
          if(suportMaxWidth && this.width > suportMaxWidth){ msj += "El ancho de <strong>"  + this.width + "px</strong> supera el máximo permitido que es: <strong>" + suportMaxWidth + "px</strong><br>"; } 
          if(suportMaxHeight && this.height > suportMaxHeight){ msj += "El alto de <strong>"  + this.height + "px</strong> supera el máximo permitido que es: <strong>" + suportMaxHeight + "px</strong><br>"; }
          if(suportMinWidth && this.width < suportMinWidth){ msj += "El ancho de <strong>"  + this.width + "px</strong> es menor al mínimo permitido que es: <strong>" + suportMinWidth + "px</strong><br>"; } 
          if(suportMinHeight && this.height < suportMinHeight){ msj += "El alto de <strong>"  + this.height + "px</strong> es menor al mínimo permitido que es: <strong>" + suportMinHeight + "px</strong>"; }


          if(msj != ''){

            msj = "<span>ERROR:</span>" + msj;
            uploadNew.msjScreen.show(e, msj, 'error');
            _URL.revokeObjectURL(objectUrl);

            return false;

          }else{

            //return true;
            //Is valid dimension > Init upload!
            uploadNew.upload(e, wrapperFile);

          } 

        };

        img.src = objectUrl;

    }

  },

  //Return extension file
  fileExtension : function(fileName){

    return fileName.substring(fileName.lastIndexOf('.')+1, fileName.length) || fileName;

  },

  //Return true if is image file
  fileIsImage : function(fileName){

    let imgSuport = new Array("jpg", "jpeg", "png", "gif", "bmp", "webp"),
        fExt = uploadNew.fileExtension(fileName),
        isImage = false;

    if(imgSuport.indexOf(fExt) != -1) isImage = true;

    return isImage;

  },

  //The umpload method
  upload : function(e, wrapperFile) {

    console.log(wrapperFile);
    //Set vars/const
    const node_prgressBarWrapper = wrapperFile.querySelector('.loaderbar-progress'),
          node_prgressBar = wrapperFile.querySelector('.loaderbar-progress .progress-bar'),
          node_previewWrapper = wrapperFile.querySelector('.upload--file-data .upload--preview-wrapper'),
          node_preview = wrapperFile.querySelector('.upload--file-data .upload--preview-wrapper .upload--preview'),
          node_deleteFile = wrapperFile.querySelector('.upload--delete-file'),
          hidden_file = wrapperFile.querySelector('.upload--file-save');

    const file_type = e.getAttribute('data-file-type') ? e.getAttribute('data-file-type') : '',
          file_path = e.getAttribute('data-file-path') ? e.getAttribute('data-file-path') : '';    
    //
    
    //Formdata
    var data = new FormData(),
        upfile = e.files;
    data.append("upfile", upfile[0]);

    //Extra data
    data.append("file_type", file_type);
    data.append("file_path", file_path);
    //
  
    //Ajax
    var xhr = new XMLHttpRequest();
    xhr.open("POST", uploadNew.option.uploadServerSide + "?" + new Date().getTime(), true);
  
    xhr.upload.addEventListener("progress", function(evt){

      if (evt.lengthComputable) {

        let percent = parseInt(evt.loaded / evt.total * 100);
            
        console.log(percent + "%");
  
        node_prgressBar.style.width = percent + '%';
        node_prgressBar.innerHTML = '<div id="progress-status">' + percent +' %</div>';
  
      }

    }, false);
  
    xhr.onloadstart = function (e) {
      
        console.log("start");
        node_prgressBarWrapper.style.display = 'block';

    }
  
    xhr.onloadend = function (e) {

        console.log("end");
        
    }
  
    xhr.onload = function(res){

      var response = JSON.parse(this.response);

      node_prgressBarWrapper.style.display = 'none';

      if(response.error == '') {
        
        // @TODO - OK
        node_prgressBar.style.width = '0%';
        node_prgressBar.innerHTML = '';
        //node_preview.setAttribute('src', '');
  
        //Si es una imágen
        if(response.file_type == 'image'){

          node_previewWrapper.classList.add('loading-file');
  
          let img = new Image();
          img.src = response.preview;
          img.addEventListener("load", function(e) {
            node_preview.setAttribute('src', response.preview);
            node_previewWrapper.classList.remove('loading-file');
            node_deleteFile.classList.remove('disabled');

            //Set preview modal
            uploadNew.setFullPreview(node_previewWrapper, response.record_saved, response.file_type);
            //

          }); 
        
        //Es un archivo (response.file_type 'doc')
        }else{

          node_previewWrapper.classList.remove('loading-file');
          node_deleteFile.classList.remove('disabled');
          node_preview.setAttribute('src', uploadNew.option.previewDocIcon + response.extention + '.svg');

          //Set preview link
          uploadNew.setFullPreview(node_previewWrapper, response.record_saved, response.file_type);
          //
 
        }

        //String con el registro a guardar en DDBB
        hidden_file.value = response.record_saved;
        //
        
        console.log("OK");
        console.log(response);
  
      } else {
        //Error
        console.log(response);

        uploadNew.deleteFile(node_deleteFile, false);
        uploadNew.msjScreen.show(e, "<span>ERROR:</span>" + response.message, 'error');
      }

    };
  
    //Send Ajax
    xhr.send(data);
    return false;
  
  },

  //Set UI onload
  ui : {

    set : function(e, i){

      uploadNew.ui.htmlPreview(e);
      uploadNew.modal.init(e, i);

    },

    htmlPreview : function(e){

      let wrapper = e.parentNode.parentNode.querySelector('.upload--preview-wrapper'),
          linkNode = e.parentNode.parentNode.querySelector('.upload--preview-wrapper a'),
          previewNode = e.parentNode.parentNode.querySelector('.upload--preview-wrapper .upload--preview'),
          uploadType = e.getAttribute('data-file-type');
  
      //Add class names to preview wrapper
      if(linkNode.getAttribute('data-uploaded-full-view') && linkNode.getAttribute('data-uploaded-full-view').length > 3){
  
        wrapper.classList.add('upload--preview-link', uploadType);
  
      }else{

        let imgType = (uploadType == 'image') ? uploadNew.option.noImage : uploadNew.option.noFile;  
        previewNode.setAttribute('src', imgType);

      }

    }
 
  },

  //Delete preview file
  deleteFile : function(e, event){

    const node_wrapper = e.parentNode.querySelector('.upload--file-data .upload--preview-wrapper'),
          node_preview = e.parentNode.querySelector('.upload--file-data .upload--preview-wrapper .upload--preview'),
          hidden_file = e.parentNode.querySelector('.upload--file-save'),
          node_file = e.parentNode.querySelector('.fileinput-button .upload--file');

    if(event && !e.classList.contains('disabled')) {

      let defaultImage = node_file.getAttribute('data-file-type') ? node_file.getAttribute('data-file-type') == 'image' ? uploadNew.option.noImage : uploadNew.option.noFile : uploadNew.option.noFile; 

      node_preview.setAttribute('src', defaultImage);
      hidden_file.value = '';
      node_file.value = '';

      e.classList.add('disabled');
      node_wrapper.classList.remove('upload--preview-link', 'image', 'file');

    }else{

      hidden_file.value = '';
      node_file.value = '';

    }

  },

  setFullPreview : function(wrap, url, type){

    let previewLink = wrap.querySelector('a'),
        modalImage = wrap.parentNode.parentNode.querySelector('.upload--modal-img-center');

    wrap.classList.add('upload--preview-link', type);

    if(type == 'doc'){

      previewLink.setAttribute('href', '../' + url);

    }else{

      //Set image in modal node
      modalImage.setAttribute('src', '../' + url);

    }

  },

  msjScreen : {

    show : function(e, msj, type){

      const target = e.parentNode.parentNode.querySelector('.upload--file-data');
      
      const msjContainer = document.createElement('div');
      msjContainer.className = 'upload--msj ' + type;

      const msjText = document.createElement('div');
      msjText.className = 'upload--msj-text';
      msjText.innerHTML = msj + '<i class="fa fa-times upload--msj-cose" aria-hidden="true"></i>';

      target.parentNode.insertBefore(msjContainer, target.nextSibling);
      msjContainer.insertBefore(msjText, null);

    },

    delete : function(e){

      const par = e.parentNode.parentNode,
            chi = par.querySelector('.upload--msj');

      if(chi) par.removeChild(chi);

    },

    events : function(){

      let wrapper = document.getElementsByClassName("upload--wrapper");

      Array.prototype.forEach.call(wrapper, function(element, index){

        element.addEventListener('click',function(e){

          if(e.target.classList.contains('upload--msj-cose')) {

            this.removeChild(this.querySelector('.upload--msj'));

          }
        
        });

      }); 

    }

  },

  //Modal window
  modal : {

    option : {

      modalActiveClass : "modal-active"

    },

    init : function(e, i){

      if(uploadNew.modal.isImage(e)){

        let linkImage = e.parentNode.parentNode.querySelector('.upload--preview-wrapper a'),
        srcImage = '';

        //Si hay imegen cargada previamente se setea la propiedad src en el modal
        if(linkImage.getAttribute('data-uploaded-full-view') && linkImage.getAttribute('data-uploaded-full-view').length > 3){

          srcImage = 'src="' + linkImage.getAttribute('data-uploaded-full-view') + '"';

        }

        let wrapper = e.parentNode.parentNode,
            modalHtml = document.createElement('div');
            modalHtml.className = 'modal modal--fullscreen';
            modalHtml.innerHTML = '<div class="upload--modal-dialog">' +
                                    '<button class="upload--modal-close" data-modal-dismiss><i class="fa fa-times" aria-hidden="true"></i></button>' +
                                    '<img class="upload--modal-img-center"' + srcImage + ' />' +
                                  '</div>';

        modalHtml.setAttribute('data-modal-name', 'sample-modal-' + i);

        wrapper.insertBefore(modalHtml, null);

        //Set link
        let previewLink = wrapper.querySelector('.upload--preview-wrapper a');

        previewLink.setAttribute('data-modal-trigger', 'sample-modal-' + i); 
        //

      }

    },

    isImage : function(e){

      let status = (e.getAttribute('data-file-type') && e.getAttribute('data-file-type') == 'doc') ? false : true;
      return status;

    },

    events : function(){

      let btnOpen = document.querySelectorAll('[data-modal-trigger]'),
          btnClose = document.querySelectorAll('[data-modal-dismiss]');

      Array.prototype.forEach.call(btnOpen, function(element, index){

        element.addEventListener('click',function(ev){

          if(ev) ev.preventDefault();

          let modalData = element.getAttribute('data-modal-trigger'),
          modalNode = document.querySelector('[data-modal-name="'+ modalData + '"]');

          modalNode.classList.add(uploadNew.modal.option.modalActiveClass);
        
        });

      });

      Array.prototype.forEach.call(btnClose, function(element, index){

        element.addEventListener('click',function(ev){

          if(ev) ev.preventDefault();

          element.parentNode.parentNode.classList.remove(uploadNew.modal.option.modalActiveClass);
        
        });

      });


    },

  }

}

