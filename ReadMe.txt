            
��ԊȒP�ȍ\��
	    const String SHEET_SOFT_LIST = "SoftwareList";
            
            /*** �����̃u�b�N�A�����̃V�[�g��ǂޏꍇ ***/
            var book = new WorkBook(@".\SoftwareManagement.xlsx");
            book.AddSheet(new WorkSheet(SHEET_SOFT_LIST));

            // �V�[�g�Ǎ���
            book.ReadBook(SHEET_SOFT_LIST);

            // �f�[�^������Ƃ���
            var list = new String[] { "notepad", "paint", "calculator" };

            // �Q�ƃZ���쐬
            var currentCell = new Range("B2");

            foreach (var item in list)  // B2 �` B4�ɒl����������
            {
                book[SHEET_SOFT_LIST][currentCell] = item;
                currentCell.Shift(0, 1);   // ����ֈړ�
            }

            // �ύX��ۑ�
            book.WriteBook();


�����̃����������ʂ����炷�H�v
        static void Main(string[] args)
        {
            /*** �����̃u�b�N�A�����̃V�[�g��ǂޏꍇ ***/
            var book = new WorkBook(@".\SoftwareManagement.xlsx");
            book.AddSheet(new SoftListSheet());

            // �V�[�g�Ǎ���
            book.ReadBook(SheetName.SOFT_LIST);

            // �f�[�^������Ƃ���
            var list = new String[] { "notepad", "paint", "calculator" };

            // �Q�ƃZ���쐬
            var currentCell = new Range("B2");

            foreach (var item in list)  // B2 �` B4�ɒl����������
            {
                book[SheetName.SOFT_LIST][currentCell] = item;
                currentCell.Shift(0, 1);   // ����ֈړ�
            }

            // �ύX��ۑ�
            book.WriteBook();
        }

        // ���ʂ̎Q��
        class SheetName
        {
            public const String SOFT_LIST = "SoftwareList";
        }
    
        // ��`���̏�Ԃ���1000 x 1000�͈̔͂�ݒ肷��̂ł��̒�`�ŃR�X�g��ߖ񂷂�
        [ExcelSheet(EnumSheetPermission.ReadWrite, ColMax = 20, RowMax = 20)]
        class SoftListSheet : WorkSheet
        {
            public SoftListSheet() : base(SheetName.SOFT_LIST) { }
        }