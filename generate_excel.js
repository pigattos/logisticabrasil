const XLSX = require('xlsx');

const geral = [
    { id: 'NC001', data: '2024-05-20', pedido: '1001', codigo: 'PROD-A', vendedor: 'Ana', setor: 'Comercial' },
    { id: 'NC002', data: '2024-05-20', pedido: '1002', codigo: 'PROD-B', vendedor: 'Bruno', setor: 'Produção' },
    { id: 'NC003', data: '2024-05-21', pedido: '1003', codigo: 'PROD-C', vendedor: 'Carlos', setor: 'Comercial' },
    { id: 'NC004', data: '2024-05-21', pedido: '1004', codigo: 'PROD-D', vendedor: 'Ana', setor: 'Financeiro' },
    { id: 'NC005', data: '2024-05-22', pedido: '1005', codigo: 'PROD-E', vendedor: 'Daniela', setor: 'Produção' },
    { id: 'NC006', data: '2024-05-22', pedido: '1006', codigo: 'PROD-F', vendedor: 'Ana', setor: 'Comercial' },
    { id: 'NC007', data: '2024-05-23', pedido: '1007', codigo: 'PROD-G', vendedor: 'Bruno', setor: 'Comercial' },
    { id: 'NC008', data: '2024-05-23', pedido: '1008', codigo: 'PROD-H', vendedor: 'Carlos', setor: 'Comercial' }
];

// In this scenario, "Ana" registered NC001 after notification, but NC006 is still missing.
const pessoal = [
    { id: 'NC001', data: '2024-05-20', pedido: '1001', codigo: 'PROD-A', vendedor: 'Ana' },
    { id: 'NC002', data: '2024-05-21', pedido: '1002', codigo: 'PROD-B', vendedor: 'Bruno' }
];

const wbGeral = XLSX.utils.book_new();
const wsGeral = XLSX.utils.json_to_sheet(geral);
XLSX.utils.book_append_sheet(wbGeral, wsGeral, "Geral");
XLSX.writeFile(wbGeral, "data_geral.xlsx");

const wbPessoal = XLSX.utils.book_new();
const wsPessoal = XLSX.utils.json_to_sheet(pessoal);
XLSX.utils.book_append_sheet(wbPessoal, wsPessoal, "Pessoal");
XLSX.writeFile(wbPessoal, "data_pessoal.xlsx");

console.log("Arquivos Excel atualizados com sucesso!");
