import os
import pytest
from mdt_doc_utils.mdt_doc_utils import WordDocumentManager

@pytest.fixture
def sample_data():
    muc_data = [
        {'GMIS_MUC_ID': 'GMIS-7888-MUC-7888', 
         'MUC_Purpose': 'ABC  CDF  XYZ', 
         'Model_User': 'User1', 'MRMR': '4'}
    ]

    roles_data = [
        {'Role': 'Model Validator', 'Name': 'John Doe', 'Title': 'Validator', 'Email': 'john@example.com'},
        {'Role': 'Peer Reviewer (optional)', 'Name': '', 'Title': '', 'Email': ''},
        {'Role': 'Approver', 'Name': 'Emily Brown', 'Title': 'Approver', 'Email': 'emily@example.com'}
    ]

    mdt_data = [
        {'#': 1, 'MDT_name': '<MDT>', 'Reception_status': 'Received'}
    ]

    generic_data = [
        ['1', '<title 1>', '', '']
    ]

    return {
        'muc_data': muc_data,
        'roles_data': roles_data,
        'mdt_data': mdt_data,
        'generic_data': generic_data
    }

@pytest.fixture
def document_manager():
    template_path = 'templates/word_template.docx'
    output_path = 'output/test_populated_word_document.docx'
    manager = DocumentManager(template_path)
    yield manager
    manager.save(output_path)
    assert os.path.exists(output_path)

def test_doc_gen(document_manager, sample_data):
    muc_table = MUCInScopeTable(document_manager.doc)
    muc_table.populate(sample_data['muc_data'])

    roles_table = RolesAndResponsibilitiesTable(document_manager.doc)
    roles_table.populate(sample_data['roles_data'])

    mdt_table = MDTInfoTable(document_manager.doc)
    mdt_table.populate(sample_data['mdt_data'])

    generic_table = GenericTable(document_manager.doc, 3)
    generic_table.populate(sample_data['generic_data'])
