import * as React from 'react';
import styles from './CadastroMemorandos.module.scss';
import { ICadastroMemorandosProps } from './ICadastroMemorandosProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';
SPComponentLoader.loadCss('https://getbootstrap.com/docs/4.1/dist/css/bootstrap.min.css');
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ComboBoxListItemPicker } from '@pnp/spfx-controls-react/lib/listItemPicker';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import { IList } from "@pnp/sp/lists";
import { CKEditor } from '@ckeditor/ckeditor5-react';
import ClassicEditor from '@ckeditor/ckeditor5-build-classic';
import b64toBlob from 'b64-to-blob'; // https://www.npmjs.com/package/b64-to-blob (npm install b64-to-blob)
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";

export default class CadastroMemorandos extends React.Component<ICadastroMemorandosProps, any> {

  public constructor(props: ICadastroMemorandosProps) {
    super(props);
    this.state = {
      hideAssinaturasExercicios: false,
      CKEditorEvent: '',
      logo: {
        title: '',
        id: ''
      },
      destinatario: {
        id: '',
        text: '',
        loginName: '',
        secondaryText: ''
      },
      origem: {
        id: '',
        text: '',
        loginName: '',
        secondaryText: ''
      },
      assunto: '',
      data: '',
      notificar: [],
      vocativo: '',
      saudacaoFinal: '',
      assinaturas: [],
      assinaturasExercicio: [],
      anexos: [],
      anexoFisico: false,
      objetoSalvo: null,
      enviadoAprovacao: false,
      siglaOrigem: '',
      siglaDestino: ''
    };
  }

  private functShowModal(openModal) {
    document.getElementById("btnModalSim").style.display = 'none';
    document.getElementById("btnModalFechar").innerText = 'OK';

    if (openModal) {
      $('#myDivMain').css('background-color', '#d3d3d3');
      $('#myDivMain').css('opacity', 0.2);
      $('#myDivMain :button,input').attr('readonly', true);
      $('#myDialog').show();
    } else {
      $('#myDivMain').css('background-color', '#ffffff');
      $('#myDivMain').css('opacity', 1);
      $('#myDivMain :button,input').attr('readonly', false);
      $('#myDialog').hide();
    }
  }

  private saveItem = () => {
    $('.modal-body').html('Salvando...');
    document.getElementById("btnModalFechar").innerText = 'OK';
    document.getElementById("selectEmpresa").focus();
    this.functShowModal(true);

    let htmlSalvar = this.state.CKEditorEvent.replace(/<br>/g, '<br/>');
    htmlSalvar = '<p style="text-align: justify">' + htmlSalvar + '</p>';

    let codigoMatriculaOrigem = this.state.origem.loginName.split("i:0#.f|membership|")[1].split("@")[0];
    console.log("origem: " + codigoMatriculaOrigem);

    let valorOrigem = '';
    let urlOrigem = 'https://celshpnth.celesc.com.br:1002/celesc/api/ad/GetSiglaLocacao/?pMatricula=' + codigoMatriculaOrigem;

    $.ajax({
      url: urlOrigem,
      type: "GET",
      success: (resultDataOrigem) => {
        this.setState({ siglaOrigem: resultDataOrigem });

        let codigoMatriculaDestino = this.state.destinatario.loginName.split("i:0#.f|membership|")[1].split("@")[0];
        console.log("destino: " + codigoMatriculaDestino);

        let valorDestino = '';
        let urlDestino = 'https://celshpnth.celesc.com.br:1002/celesc/api/ad/GetSiglaLocacao/?pMatricula=' + codigoMatriculaDestino;

        $.ajax({
          url: urlDestino,
          type: "GET",
          success: (resultDataDestino) => {
            this.setState({ siglaDestino: resultDataDestino });

            let objSalvar = {
              texto: htmlSalvar,
              destinatario: this.state.destinatario,
              logo: this.state.logo,
              origem: this.state.origem,
              assunto: this.state.assunto,
              data: this.state.data,
              notificar: { 'results': this.state.notificar },
              vocativo: this.state.vocativo,
              saudacaoFinal: this.state.saudacaoFinal,
              assinaturas: { 'results': this.state.assinaturas },
              assinaturasExercicio: { 'results': this.state.assinaturasExercicio },
              anexos: this.state.anexos,
              anexoFisico: this.state.anexoFisico,
              id: 0,
              UnidadeOrigem: this.state.siglaOrigem,
              UnidadeDestino: this.state.siglaDestino
            }

            console.log(objSalvar);

            let camposNaoInformados = '';

            if (objSalvar.logo.id == 0) {
              camposNaoInformados += '<br>- Empresa';
            }

            if (objSalvar.origem.id == "") {
              camposNaoInformados += '<br>- Origem';
            }

            if (objSalvar.data == "") {
              camposNaoInformados += '<br>- Data';
            }

            if (objSalvar.destinatario.id == "") {
              camposNaoInformados += '<br>- Destinatário';
            }

            if (objSalvar.vocativo == "") {
              camposNaoInformados += '<br>- Vocativo';
            }

            if (objSalvar.assunto == "") {
              camposNaoInformados += '<br>- Assunto';
            }

            if (objSalvar.texto == '') {
              camposNaoInformados += '<br>- Corpo Texto';
            }

            if (objSalvar.saudacaoFinal == "") {
              camposNaoInformados += '<br>- Saudação Final';
            }

            if (objSalvar.assinaturas.results.length < 1 && objSalvar.assinaturasExercicio.results.length < 1) {
              camposNaoInformados += '<br>- Assinaturas ou Assinaturas em Exercício';
            }

            if (camposNaoInformados != '') {
              //alert('Campos não informados: ' + camposNaoInformados);
              $('.modal-body').html('Campos não informados: ' + camposNaoInformados);
              document.getElementById("selectEmpresa").focus();
              this.functShowModal(true);
            } else {

              sp.setup({
                spfxContext: this.props.context
              });

              console.log('enviado para aprovação: ' + this.state.enviadoAprovacao);
              console.log('salvo: ' + this.state.objetoSalvo);
              // if true é a primeira vez que o item é criado, nas próximas será editado
              if (this.state.objetoSalvo == null) {
                try {
                  sp.web.lists.getByTitle('Memorandos').items.add({
                    Title: objSalvar.assunto,
                    EmpresaId: objSalvar.logo.id,
                    OrigemId: objSalvar.origem.id,
                    DestinoId: objSalvar.destinatario.id,
                    Data: objSalvar.data,
                    //MEMO: '',
                    NotificarId: objSalvar.notificar,
                    Vocativo: objSalvar.vocativo,
                    CorpoTexto: objSalvar.texto,
                    AnexoFisico: objSalvar.anexoFisico,
                    SaudacaoFinal: objSalvar.saudacaoFinal,
                    AssinaturasId: objSalvar.assinaturas,
                    AssinaturasEmExercicioId: objSalvar.assinaturasExercicio,
                    UnidadeOrigem: objSalvar.UnidadeOrigem,
                    UnidadeDestino: objSalvar.UnidadeDestino
                  }).then((result: any) => {
                    objSalvar.id = result.data.Id;
                    // salvando anexos
                    const list: IList = sp.web.lists.getByTitle("Memorandos");
                    let fileInfos: IAttachmentFileInfo[] = [];

                    this.state.anexos.map(item => {
                      fileInfos.push({
                        name: item.name,
                        content: item.content
                      });
                    });

                    try {
                      list.items.getById(result.data.Id).attachmentFiles.addMultiple(fileInfos);
                      // geração de pdf

                      try {
                        let valorOrigem = objSalvar.UnidadeOrigem;
                        let valorDestino = objSalvar.UnidadeDestino;

                        if (valorOrigem == null || valorDestino == null) {
                          console.log('destino ou origem não encontrados:'
                            + '/n+'
                            + 'Origem: ' + valorOrigem
                            + '/n+'
                            + 'Destino: ' + valorDestino);
                        }

                        let dataToAjax = {
                          "Company": objSalvar.logo.title,
                          "Source": valorOrigem != null ? valorOrigem : '',
                          "Destiny": valorDestino != null ? valorDestino : '',
                          "Subject": objSalvar.assunto,
                          "Date": objSalvar.data.split('-')[2] + '/' + objSalvar.data.split('-')[1] + '/' + objSalvar.data.split('-')[0],
                          "Title": objSalvar.vocativo,
                          "MemoBody": objSalvar.texto,
                          "Salutation": objSalvar.saudacaoFinal,
                          "Signatures": [],
                          "GeneratePhysicalFile": false
                        }

                        console.log(dataToAjax);

                        let base64content = { Status: '', File: '' };

                        $.ajax({
                          type: 'POST',
                          url: 'https://celshpnth.celesc.com.br:1002/celesc/api/memorando/createpdf',
                          data: dataToAjax,
                          dataType: "text",
                          success: (resuldivata) => {
                            base64content = JSON.parse(resuldivata);
                            if (base64content.Status == 'success') {
                              // Convertendo para Blob
                              var contentType = "application/pdf";
                              var blob = b64toBlob(base64content.File, contentType);
                              var blobUrl = URL.createObjectURL(blob);

                              // abrindo pdf
                              window.open(blobUrl, "_blank");
                              this.setState({ objetoSalvo: objSalvar });

                              //mudando o modal
                              document.getElementById("btnModalFechar").innerText = 'Não';
                              document.getElementById("btnModalSim").style.display = 'block';
                              $('.modal-body').html('Memorando salvo com sucesso.<br><br>PDF criado, veja a aba recém aberta.<br><br>Deseja enviar o memorando para aprovação?');
                            } else {
                              console.log(base64content);
                              $('.modal-body').html('Erro na criação de PDF!');
                            }
                          }, error: (jqXHR, textStatus, errorThrown) => {
                            console.log("Erro na API de criação de PDF");
                            $('.modal-body').html('Erro na API de criação de PDF!');
                          }
                        });
                      } catch (e) {
                        console.log("Erro na criação de PDF");
                        $('.modal-body').html('Erro na criação de PDF!');
                        console.log(e);
                      }
                    } catch (e) {
                      console.log("Erro na adição de anexo");
                      $('.modal-body').html('Erro na adição de anexo!');
                      console.log(e);
                    }
                  });
                } catch (e) {
                  console.log("Erro na criação do item na lista de memorandos");
                  $('.modal-body').html('Erro na criação do item na lista de memorandos!');
                  console.log(e);
                }
              } else if (this.state.enviadoAprovacao != true) {
                let list = sp.web.lists.getByTitle("Memorandos");
                try {
                  list.items.getById(this.state.objetoSalvo.id).update({
                    Title: objSalvar.assunto,
                    EmpresaId: objSalvar.logo.id,
                    OrigemId: objSalvar.origem.id,
                    DestinoId: objSalvar.destinatario.id,
                    Data: objSalvar.data,
                    //MEMO: '',
                    NotificarId: objSalvar.notificar,
                    Vocativo: objSalvar.vocativo,
                    CorpoTexto: objSalvar.texto,
                    AnexoFisico: objSalvar.anexoFisico,
                    SaudacaoFinal: objSalvar.saudacaoFinal,
                    AssinaturasId: objSalvar.assinaturas,
                    AssinaturasEmExercicioId: objSalvar.assinaturasExercicio,
                    UnidadeOrigem: objSalvar.UnidadeOrigem,
                    UnidadeDestino: objSalvar.UnidadeDestino
                  }).then((result: any) => {
                    objSalvar.id = this.state.objetoSalvo.id;
                    // salvando anexos
                    const list: IList = sp.web.lists.getByTitle("Memorandos");
                    let fileInfos: IAttachmentFileInfo[] = [];

                    this.state.anexos.map(item => {
                      fileInfos.push({
                        name: item.name,
                        content: item.content
                      });
                    });
                    //console.log(fileInfos);
                    try {
                      let idItemSalvo = this.state.objetoSalvo.id;
                      // get all the attachments
                      list.items.getById(this.state.objetoSalvo.id).attachmentFiles().then(anexos => {

                        let delecaoAnexos = new Promise(async (resolve, reject) => {
                          for (var i = 0; anexos.length > i; i++) {
                            console.log('Excluindo arquivo: ' + anexos[i].FileName);
                            await list.items.getById(idItemSalvo).attachmentFiles.deleteMultiple(anexos[i].FileName);
                          }
                          resolve(true);
                        });

                        delecaoAnexos.then(result => {
                          console.log('Adicionando novos anexos');
                          list.items.getById(idItemSalvo).attachmentFiles.addMultiple(fileInfos);
                          // geração de pdf

                          try {
                            let valorOrigem = objSalvar.UnidadeOrigem;
                            let valorDestino = objSalvar.UnidadeDestino;

                            if (valorOrigem == null || valorDestino == null) {
                              console.log('destino ou origem não encontrados:'
                                + '/n+'
                                + 'Origem: ' + valorOrigem
                                + '/n+'
                                + 'Destino: ' + valorDestino);
                            }

                            let dataToAjax = {
                              "Company": objSalvar.logo.title,
                              "Source": valorOrigem != null ? valorOrigem : '',
                              "Destiny": valorDestino != null ? valorDestino : '',
                              "Subject": objSalvar.assunto,
                              "Date": objSalvar.data.split('-')[2] + '/' + objSalvar.data.split('-')[1] + '/' + objSalvar.data.split('-')[0],
                              "Title": objSalvar.vocativo,
                              "MemoBody": objSalvar.texto,
                              "Salutation": objSalvar.saudacaoFinal,
                              "Signatures": [],
                              "GeneratePhysicalFile": false
                            }

                            console.log(dataToAjax);

                            let base64content = { Status: '', File: '' };

                            $.ajax({
                              type: 'POST',
                              url: 'https://celshpnth.celesc.com.br:1002/celesc/api/memorando/createpdf',
                              data: dataToAjax,
                              dataType: "text",
                              success: (resuldivata) => {
                                base64content = JSON.parse(resuldivata);
                                if (base64content.Status == 'success') {
                                  // Convertendo para Blob
                                  var contentType = "application/pdf";
                                  var blob = b64toBlob(base64content.File, contentType);
                                  var blobUrl = URL.createObjectURL(blob);

                                  // abrindo pdf
                                  window.open(blobUrl, "_blank");
                                  this.setState({ objetoSalvo: objSalvar });

                                  //mudando o modal
                                  document.getElementById("btnModalFechar").innerText = 'Não';
                                  document.getElementById("btnModalSim").style.display = 'block';
                                  $('.modal-body').html('Memorando salvo com sucesso.<br><br>PDF criado, veja a aba recém aberta.<br><br>Deseja enviar o memorando para aprovação?');
                                } else {
                                  console.log(base64content);
                                  $('.modal-body').html('Erro na criação de PDF!');
                                }
                              }, error: (jqXHR, textStatus, errorThrown) => {
                                console.log("Erro na API de criação de PDF");
                                $('.modal-body').html('Erro na criação de PDF!');
                              }
                            });
                          } catch (e) {
                            console.log("Erro na criação de PDF");
                            $('.modal-body').html('Erro na criação de PDF!');
                            console.log(e);
                          }
                        }).catch(err => {
                          console.log(err);
                        });
                      });
                    } catch (e) {
                      console.log("Erro na adição de anexo");
                      $('.modal-body').html('Erro na adição de anexo!');
                      console.log(e);
                    }
                  });
                } catch (e) {
                  console.log("Erro na criação do item na lista de memorandos");
                  $('.modal-body').html('Erro na criação do memorando!');
                  console.log(e);
                }
              } else if (this.state.enviadoAprovacao == true) {
                $('.modal-body').html('O item foi enviado para aprovação, não pode mais ser editado!');
              }
            }
          }, error: (jqXHR, textStatus, errorThrown) => {
            console.log("Erro na API de busca de Sigla Locação de destino");
            $('.modal-body').html('Erro na API de busca de Sigla Locação de destino!');
          }
        });

      }, error: (jqXHR, textStatus, errorThrown) => {
        console.log("Erro na API de busca de Sigla Locação de origem");
        $('.modal-body').html('Erro na API de busca de Sigla Locação de origem!');
      }
    });

    // Montando objeto a ser salvo

  }

  public salvarAssinaturas = () => {
    document.getElementById("btnModalSim").style.display = 'none';
    document.getElementById("btnModalFechar").innerText = 'OK';

    $('.modal-body').html('Salvando assinaturas...');
    console.log(this.state.objetoSalvo);

    if (this.state.objetoSalvo.assinaturasExercicio.results.length > 0) {
      this.state.objetoSalvo.assinaturasExercicio.results.map(aprovador => {
        try {
          //salvando assinatura
          sp.setup({
            spfxContext: this.props.context
          });

          let assinatura = {
            titulo: "Assinatura do memorando de Assunto: " + this.state.objetoSalvo.assunto,
            ListaDeOrigem: 'Memorandos',
            DocumentoRelacionado: this.state.objetoSalvo.id,
            Aprovador: aprovador
          }

          sp.web.lists.getByTitle('Assinaturas').items.add({
            Title: assinatura.titulo,
            ListaDeOrigem: assinatura.ListaDeOrigem,
            DocumentoRelacionado: assinatura.DocumentoRelacionado,
            AprovadorId: assinatura.Aprovador
          }).then((result: any) => {
            this.setState({ enviadoAprovacao: true });
            $('.modal-body').html('Memorando enviado para aprovação com sucesso!');
          });

          let list = sp.web.lists.getByTitle("Memorandos");
          list.items.getById(this.state.objetoSalvo.id).update({
            SubmetidoParaAprovacao: true
          });

        } catch (e) {
          console.log("Erro na criação do item na lista de assinaturas");
          $('.modal-body').html('Não foi possível enviar para aprovação!');
          console.log(e);
        }
      });
    }

    if (this.state.objetoSalvo.assinaturas.results.length > 0) {
      this.state.objetoSalvo.assinaturas.results.map(aprovador => {
        try {
          //salvando assinatura
          sp.setup({
            spfxContext: this.props.context
          });

          let assinatura = {
            titulo: "Assinatura do memorando de Assunto: " + this.state.objetoSalvo.assunto,
            ListaDeOrigem: 'Memorandos',
            DocumentoRelacionado: this.state.objetoSalvo.id,
            Aprovador: aprovador
          }

          sp.web.lists.getByTitle('Assinaturas').items.add({
            Title: assinatura.titulo,
            ListaDeOrigem: assinatura.ListaDeOrigem,
            DocumentoRelacionado: assinatura.DocumentoRelacionado,
            AprovadorId: assinatura.Aprovador
          }).then((result: any) => {
            this.setState({ enviadoAprovacao: true });
            $('.modal-body').html('Memorando enviado para aprovação com sucesso!');
          });

          let list = sp.web.lists.getByTitle("Memorandos");
          list.items.getById(this.state.objetoSalvo.id).update({
            SubmetidoParaAprovacao: true
          });
        } catch (e) {
          console.log("Erro na criação do item na lista de assinaturas");
          $('.modal-body').html('Não foi possível enviar para aprovação!');
          console.log(e);
        }
      });
    }
  }

  public componentDidMount() {
    this.mostraEscondeAssinaturasExercicio();

    // Modal - Dialog
    $('#myDialog').css('border', 2);
    $('#myDialog').css('border-style', 'solid');
    $('#myDialog').css('border-radius', 5);
  }

  public onSelectedComboBoxItemPicker = (evt: any) => {
    let temp = '';
    switch (evt.target.options[evt.target.selectedIndex].text) {
      case 'Celesc Holding':
        temp = 'Holding';
        break;
      case 'Celesc Distribuição S.A.':
        temp = 'Distribuição';
        break;
      case 'Celesc Geração S.A.':
        temp = 'Geração';
        break;
      default:
        break;
    }

    this.setState({
      logo: {
        title: temp,
        id: evt.target.selectedIndex
      }
    });
  }

  public GetPeoplePickerNotificar = (items: any[]) => {
    let arrayTemp = [];
    items.map((item) => {
      arrayTemp.push(
        item.id
        //text: item.text,
        //loginName: item.loginName,
        //secondaryText: item.secondaryText
      )
    });
    this.setState({ notificar: arrayTemp });
  }

  public GetPeoplePickerAssinaturas = (items: any[]) => {
    let arrayTemp = [];
    items.map((item) => {
      arrayTemp.push(
        item.id
        //text: item.text,
        //loginName: item.loginName,
        //secondaryText: item.secondaryText
      )
    });
    this.setState({ assinaturas: arrayTemp });
  }

  public GetPeoplePickerAssinaturasExercicio = (items: any[]) => {
    let arrayTemp = [];
    items.map((item) => {
      arrayTemp.push(
        item.id
        //text: item.text,
        //loginName: item.loginName,
        //secondaryText: item.secondaryText
      )
    });
    this.setState({ assinaturasExercicio: arrayTemp });
  }

  public GetPeoplePickerDestinatario = (items: any[]) => {
    items.map((item) => {
      this.setState({
        destinatario: {
          id: item.id,
          text: item.text,
          loginName: item.loginName,
          secondaryText: item.secondaryText
        }
      });
    });
  }

  public GetPeoplePickerOrigem = (items: any[]) => {
    items.map((item) => {
      this.setState({
        origem: {
          id: item.id,
          text: item.text,
          loginName: item.loginName,
          secondaryText: item.secondaryText
        }
      });
    });
  }

  public GetAssunto = (evt: any) => {
    this.setState({
      assunto: evt.target.value
    });
  }

  public GetVocativo = (evt: any) => {
    this.setState({
      vocativo: evt.target.value
    });
  }

  public GetSaudacaoFinal = (evt: any) => {
    this.setState({
      saudacaoFinal: evt.target.value
    });
  }

  public GetAnexoFisico = (evt: any) => {
    this.setState({
      anexoFisico: evt.target.checked
    });
  }

  public getDateCalendario = () => {
    this.setState({ data: $('#data').val() });
  }

  public mostraEscondeAssinaturasExercicio = () => {
    //show/hide tr
    if (this.state.hideAssinaturasExercicios) {
      $('#trAssinaturasExercicio').show();
      $('#showHideButton').text('Fechar assinaturas em exercício');
    } else {
      $('#trAssinaturasExercicio').hide();
      $('#showHideButton').text('Adicionar assinaturas em exercício');
      this.setState({ assinaturasExercicio: [] });
    }

    //setando novo state
    let temp = this.state.hideAssinaturasExercicios ? false : true;
    this.setState({ hideAssinaturasExercicios: temp });
  }

  public return = () => {
    window.location.href = 'https://celesccombr.sharepoint.com/sites/ecm/SitePages/Memorandos.aspx';
  }

  public setAnexo(item) {
    // obtendo itens já inclusos
    let tempArray = this.state.anexos;
    //lendo o arquivo
    item.downloadFileContent().then(r => {
      let reader = new FileReader();
      reader.readAsArrayBuffer(r);
      reader.onload = () => {
        //adicionando o arquivo lido no array
        tempArray.push({
          name: item.fileName,
          content: reader.result
        });
        //adicionando array no state
        this.setState({ anexos: tempArray });
      }
    });
  }

  public removeAtachArray(i) {
    let arrayTemp = this.state.anexos;
    arrayTemp.splice(i, 1);
    this.setState({ anexos: arrayTemp });
  }


  public render(): React.ReactElement<ICadastroMemorandosProps> {
    return (
      <div className={styles.testeMemorando} id="mainScrollDiv">
        <div className={styles.container} id="myDivMain">
          <div className={styles.row}>
            <div>
              <div className={styles.tabelaGeral}>
                <div className="row">
                  <div className={styles.titlesMemorandos}>
                    <span>Cadastro de Memorandos teste 22</span>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Empresa</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <select className="form-control form-control-sm" id="selectEmpresa" onChange={this.onSelectedComboBoxItemPicker}>
                      <option selected value="0">Informar...</option>
                      <option value="1">Celesc Holding</option>
                      <option value="2">Celesc Distribuição S.A.</option>
                      <option value="3">Celesc Geração S.A.</option>
                    </select>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Origem</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <PeoplePicker personSelectionLimit={1} peoplePickerCntrlclassName={styles.inputPNPControls} context={this.props.context} selectedItems={this.GetPeoplePickerOrigem} ensureUser={true}></PeoplePicker>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Data</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <div className={styles.inputPNPControls}>
                      <input type="date" onChange={this.getDateCalendario} className="form-control" id="data" placeholder="" autoComplete="off" required />
                    </div>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Controle</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <span>Memorando</span>
                  </div>
                </div>
                <hr className={styles.hrMemorandos} />
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Destinatário</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <PeoplePicker personSelectionLimit={1} peoplePickerCntrlclassName={styles.inputPNPControls} context={this.props.context} selectedItems={this.GetPeoplePickerDestinatario} ensureUser={true}></PeoplePicker>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Notificar para</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <PeoplePicker personSelectionLimit={999} peoplePickerCntrlclassName={styles.inputPNPControls} context={this.props.context} selectedItems={this.GetPeoplePickerNotificar} ensureUser={true}></PeoplePicker>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Vocativo</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <input type="text" className="form-control" onChange={this.GetVocativo}></input>
                  </div>
                </div>
                <div className="row">
                  <div className={styles.titlesMemorandos}>
                    <span>Detalhes do Memorando</span>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Assunto</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <input type="text" className="form-control" onChange={this.GetAssunto}></input>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Corpo Texto</span>
                  </div>
                  <div className="col-md-9 mb-2">
                    <CKEditor editor={ClassicEditor} data=""
                      config={{
                        toolbar: {
                          items: [
                            //'heading',
                            //'|',
                            'bold',
                            'italic',
                            //'fontSize',
                            //'fontFamily',
                            //'fontColor',
                            //'fontBackgroundColor',
                            'link',
                            'bulletedList',
                            'numberedList',
                            //'imageUpload',
                            'insertTable',
                            //'blockQuote',
                            'undo',
                            'redo'
                          ]
                        },
                        image: {
                          toolbar: [
                            'imageStyle:full',
                            'imageStyle:side',
                            '|',
                            'imageTextAlternative'
                          ]
                        },
                        fontFamily: {
                          options: [
                            'Times New Roman'
                          ]
                        },
                        language: 'pt-br'
                      }}
                      onReady={editor => {
                        // You can store the "editor" and use when it is needed.
                        editor.editing.view.change(writer => {
                          writer.setStyle(
                            "height",
                            "200px",
                            editor.editing.view.document.getRoot()
                          );
                        });

                        editor.editing.view.change(writer => {
                          writer.setStyle(
                            "width",
                            "100%",
                            editor.editing.view.document.getRoot()
                          );
                        });

                        editor.editing.view.change(writer => {
                          writer.setStyle(
                            "color",
                            "black",
                            editor.editing.view.document.getRoot()
                          );
                        });
                        //console.log('Editor is ready to use!', editor);
                      }}
                      onChange={(event, editor) => {

                      }}
                      onBlur={(event, editor) => {
                        //console.log('Blur.', editor);
                        const data = editor.getData();
                        this.setState({ CKEditorEvent: data });
                      }}
                      onFocus={(event, editor) => {
                        //console.log('Focus.', editor);
                      }}
                    />
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Saudação Final</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <input type="text" className="form-control" onChange={this.GetSaudacaoFinal}></input>
                  </div>
                </div>
                <div className="row">
                  <div className={styles.titlesMemorandos}>
                    <span>Assinaturas</span>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Assinaturas</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <PeoplePicker personSelectionLimit={999} peoplePickerCntrlclassName={styles.inputPNPControls} context={this.props.context} selectedItems={this.GetPeoplePickerAssinaturas} ensureUser={true}></PeoplePicker>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm"></div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <button id="showHideButton" className="btn btn-outline-secondary btn-sm" onClick={this.mostraEscondeAssinaturasExercicio}>Adicionar assinaturas em exercício</button>
                  </div>
                </div>
                <div className="row" id="trAssinaturasExercicio">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Em Exercício</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <PeoplePicker personSelectionLimit={999} peoplePickerCntrlclassName={styles.inputPNPControls} context={this.props.context} selectedItems={this.GetPeoplePickerAssinaturasExercicio} ensureUser={true}></PeoplePicker>
                  </div>
                </div>
                <div className="row">
                  <div className={styles.titlesMemorandos}>
                    <span>Anexos</span>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-3 input-group-sm">
                    <span>Anexar</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <FilePicker
                      bingAPIKey="<BING API KEY>"
                      accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
                      buttonLabel="Selecionar arquivo"
                      onSave={(filePickerResult: IFilePickerResult) => {
                        this.setAnexo(filePickerResult);
                      }}
                      onChanged={(filePickerResult: IFilePickerResult) => {
                        this.setAnexo(filePickerResult);
                      }}
                      context={this.props.context} />
                  </div>
                </div>
                {this.state.anexos.map((item, index) =>
                  <div className="row" key={index}>
                    <div className="col-md-4 mb-2 input-group-sm"><span className={styles.spanAnexo}>{item["name"]}</span></div>
                    <div className="col-md-4 mb-2 input-group-sm"><button className="btn btn-danger btn-sm" onClick={() => {
                      this.removeAtachArray(index);
                    }}>x</button></div>
                  </div>
                )}
                <div className="row">
                  <div className="col-md-3 mb-2 input-group-sm">
                    <span>Anexo físico</span>
                  </div>
                  <div className="col-md-6 mb-2 input-group-sm">
                    <label><input type="checkbox" className="checkbox" id="checkAnexoFisico" onChange={this.GetAnexoFisico}></input> Sim</label>
                  </div>
                </div>
                <hr className={styles.hrMemorandos} />
                <div className="row">
                  <div className="col-md-12 mb-1 input-group-sm">
                    <div className="float-right">
                      <button className="btn btn-primary btn-sm" onClick={this.saveItem}>Salvar</button>&ensp;
                      <a className="btn btn-secondary btn-sm" href="https://celesccombr.sharepoint.com/sites/ecm/SitePages/Memorandos.aspx" target="_parent" data-interception="off">Cancelar</a>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <dialog id="myDialog" className={styles.classModal}>
          <form id="myFormDialog">
            <div className="modal-header">
              <h5 className="modal-title" id="TituloModalCentralizado">Mensagem</h5>
              <button type="button" className="close" data-dismiss="modal" aria-label="Fechar" onClick={() => this.functShowModal(false)}>
                <span aria-hidden="true">&times;</span>
              </button>
            </div>
            <div className="modal-body">

            </div>
            <div className="modal-footer">
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <a href="#" id="btnModalSim" className={`btn btn-primary btn-sm ` + styles.diplayNoneBtnModal} onClick={this.salvarAssinaturas}>Sim</a>
              <a href="#" id="btnModalFechar" className="btn btn-secondary btn-sm" onClick={() => this.functShowModal(false)}>OK</a>
            </div>
          </form>
        </dialog>
      </div>
    );
  }
}