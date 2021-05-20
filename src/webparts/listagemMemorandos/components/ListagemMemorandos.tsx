import * as React from 'react';
import { IListagemMemorandosProps } from './IListagemMemorandosProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as $ from "jquery";
import * as moment from 'moment';
import DataTable from 'react-bs-datatable';
import { SPComponentLoader } from '@microsoft/sp-loader';

import styles from '../components/ListagemMemorandos.module.scss';

require('../../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../../../node_modules/bootstrap-4-required/src/css/bootstrap.css');

export default class ListagemMemorandos extends React.Component<IListagemMemorandosProps, any> {
  public constructor(props: IListagemMemorandosProps) {
    super(props);
    this.state = {
      Memorandos: [{
        texto: '',
        logo: '',
        destinatario: '',
        origem: '',
        assunto: '',
        data: '',
        notificar: [],
        vocativo: '',
        saudacaoFinal: '',
        assinaturas: [],
        assinaturasExercicio: [],
        anexos: '',
        temAnexo: '',
        anexoFisico: '',
        codigo: '',
        status: '',
        dataAprovacao: '',
        criadoPor: '',
        id: ''
      }],
      tituloPagina: ''
    };
  }

  public componentDidMount() {
    switch (this.props.description) {
      case 'MemorandosPendentes':
        this.setState({ tituloPagina: 'Pendentes' });
        this.searchListMemorandos("$filter=Status eq 'Pendente'");
        break;
      default:
        $('.modal-body').html('A listagem não foi localizada.');
        this.functShowModal(true);
        break;
    }

    $('input:text').attr('placeholder', 'Pesquisar...');

    let i = 0;
    let repeatS = true;
    let repeatB = true;
    while (repeatS || repeatB) {
      // Star - span
      if (repeatS) {
        if ($('span')[i].innerText.toLowerCase() == "show") {
          $('span')[i].innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
        }

        if ($('span')[i].innerText.toLowerCase() == "entries") {
          $('span')[i].innerHTML = "&nbsp;Itens";
          repeatS = false;
        }
      }
      // End - span

      // End - button
      if (repeatB) {
        if ($('button')[i].innerText.toLowerCase() == "first") {
          $('button')[i].innerText = "<<"
        }

        if ($('button')[i].innerText.toLowerCase() == "prev") {
          $('button')[i].innerText = "<"
        }

        if ($('button')[i].innerText.toLowerCase() == "next") {
          $('button')[i].innerText = ">"
        }

        if ($('button')[i].innerText.toLowerCase() == "last") {
          $('button')[i].innerText = ">>"
          repeatB = false;
        }
      }
      // End - button

      i++;
    }
  }

  private searchListMemorandos(filter) {
    //let url = `${this.props.siteURL}/sites/ecm/_api/web/lists/getByTitle('Memorandos')/items`;
    let url = `${this.props.siteURL}/sites/ecm/_api/web/lists/getByTitle('Memorandos')/items?${filter}&$select=CorpoTexto,Id,Empresa/Title,Destino/Title,Origem/Title,Title,Data,Notificar/Title,Vocativo,SaudacaoFinal,Assinaturas/Title,AssinaturasEmExercicio/Title,AttachmentFiles,Attachments,AnexoFisico,Codigo,Status,dataAprovacao,Author/Title&$expand=Empresa,Destino,Author,AssinaturasEmExercicio,Assinaturas,Notificar,Origem`;
    //console.log(url);
    $.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose' },
      success: (data) => {
        console.log(data);
        if (data.d.results.length > 0) {
          let tempArray = [];
          for (var i = 0; i < data.d.results.length; i++) {
            let dataCriado = data.d.results[i].Data.split('T')[0];
            let aprovacao = data.d.results[i].dataAprovacao != null ? data.d.results[i].dataAprovacao.split('T')[0] : '';

            tempArray.push({
              texto: data.d.results[i].CorpoTexto,
              logo: data.d.results[i].EmpresaId,
              destinatario: data.d.results[i].Destino.Title,
              origem: data.d.results[i].Origem.Title,
              assunto: data.d.results[i].Title,
              data: dataCriado.split('-')[2] + '-' + dataCriado.split('-')[1] + '-' + dataCriado.split('-')[0],
              notificar: data.d.results[i].Notificar.length > 1 ? data.d.results[i].Notificar.results.Title : '',
              vocativo: data.d.results[i].Vocativo,
              saudacaoFinal: data.d.results[i].SaudacaoFinal,
              assinaturas: data.d.results[i].Assinaturas.length > 1 ? data.d.results[i].Assinaturas.results.Title : null,
              assinaturasExercicio: data.d.results[i].AssinaturasEmExercicio.length > 1 ? data.d.results[i].AssinaturasEmExercicio.results.Title : null,
              anexos: data.d.results[i].AttachmentFiles,
              temAnexo: data.d.results[i].Attachments,
              anexoFisico: data.d.results[i].AnexoFisico,
              codigo: data.d.results[i].Codigo,
              status: data.d.results[i].Status,
              dataAprovacao: aprovacao != '' ? aprovacao.split('-')[2] + '-' + aprovacao.split('-')[1] + '-' + aprovacao.split('-')[0] : '',
              criadoPor: data.d.results[i].Author.Title,
              id: data.d.results[i].Id
            });
          }
          console.log(tempArray);
          this.setState({ Memorandos: tempArray });
        }
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log("Erro na API");
      }
    });
  }

  private functShowModal(openModal) {
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

  public showItemDetails = (item) => {
    alert('teste');
  };

  public render(): React.ReactElement<IListagemMemorandosProps> {
    const header = [
      { title: 'Código', prop: 'codigo', sortable: true, filterable: true },
      { title: 'Origem', prop: 'unidadeOrigem', sortable: true, filterable: true },
      { title: 'Assunto', prop: 'assunto', sortable: true, filterable: true },
      { title: 'Data entrada', prop: 'entrada', sortable: true, filterable: true },
      { title: 'Criado por', prop: 'criadoPor', sortable: true, filterable: true },
      { title: 'Ações', prop: 'actions', sortable: false, filterable: false }
    ];

    const onSortFunction = {
      date(columValue) {
        return moment(columValue);
      }
    };

    // 
    // Cria um array com os Memorandos
    // 
    const newBody = new Array<object>();

    this.state.Memorandos.forEach(element => {
      //console.log(element);
      let actions = (
        <div>
          <a onClick={this.showItemDetails}><img src="https://celesccombr.sharepoint.com/sites/ecm/SiteAssets/images/icons/iconfinder_-_Eye-Show-View-Watch-See_3844476.png"></img></a>
          <a><img src="https://celesccombr.sharepoint.com/sites/ecm/SiteAssets/images/icons/iconfinder_document_text_edit_103514.png"></img></a>
          <a><img src="https://celesccombr.sharepoint.com/sites/ecm/SiteAssets/images/icons/iconfinder_Streamline-70_185090.png"></img></a>
        </div>
      );

      newBody.push({
        codigo: element.codigo,
        unidadeOrigem: element.origem,
        assunto: element.assunto,
        entrada: element.data,
        criadoPor: element.criadoPor,
        actions: actions
      });
    });

    return (
      <div className={styles.listagemContainer}>
        <h2><b>Memorandos {this.state.tituloPagina}</b></h2>
        <div className="container mt-4">
          <DataTable
            tableHeaders={header}
            tableBody={newBody}
            rowsPerPage={25}
            rowsPerPageOption={[25, 50, 100]}
            initialSort={{ prop: 'codigo' }}
            // initialSort={{prop: 'codigo', isAscending: true}}
            onSort={onSortFunction}
          />
        </div>
        <dialog id="myDialog">
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
              <a href="#" id="btnModalFechar" className="btn btn-secondary btn-sm" onClick={() => this.functShowModal(false)}>OK</a>
            </div>
          </form>
        </dialog>
      </div>
    );
  }
}