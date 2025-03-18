#include <wx/wx.h>
#include <wx/grid.h>
#include <wx/textfile.h>
#include <wx/tokenzr.h>


class MyApp : public wxApp {
public:
	bool OnInit() override;
};

wxIMPLEMENT_APP(MyApp);

class MyFrame : public wxFrame {
public:
	
	MyFrame(const wxString& title);

private:
	wxGrid* grid;

	void OnOpen(wxCommandEvent& event);
	void OnSave(wxCommandEvent& event);
	void OnExit(wxCommandEvent& event);
	void OnCellChanged(wxGridEvent& event);

};

bool MyApp::OnInit() {
	MyFrame* frame = new MyFrame("Data Pulse");
	frame->Show(true);
	return true;
}

MyFrame::MyFrame(const wxString& title)
	: wxFrame(nullptr, wxID_ANY, title, wxDefaultPosition, wxSize(800, 600)){

	grid = new wxGrid(this, wxID_ANY, wxPoint(10, 10), wxSize(780, 550));
	grid->CreateGrid(10, 5);

	grid->SetColLabelValue(0, "A");
	grid->SetColLabelValue(1, "B");
	grid->SetColLabelValue(2, "C");
	grid->SetColLabelValue(3, "D");
	grid->SetColLabelValue(4, "E");

	grid->Bind(wxEVT_GRID_CELL_CHANGED, &MyFrame::OnCellChanged, this);

	wxMenu* menuFile = new wxMenu;
	menuFile->Append(wxID_OPEN, "Open");
	menuFile->Append(wxID_SAVE, "Save");
	menuFile->Append(wxID_EXIT, "Exit");

	wxMenuBar* menuBar = new wxMenuBar;
	menuBar->Append(menuFile, "File");

	SetMenuBar(menuBar);

	Bind(wxEVT_MENU, &MyFrame::OnOpen, this, wxID_OPEN);
	Bind(wxEVT_MENU, &MyFrame::OnSave, this, wxID_SAVE);
	Bind(wxEVT_MENU, &MyFrame::OnExit, this, wxID_EXIT);

}

void MyFrame::OnCellChanged(wxGridEvent& event) {

	int row = event.GetRow();
	int col = event.GetCol();
	wxString value = grid->GetCellValue(row, col);

	if (value.StartsWith("=SUM(") && value.EndsWith(")")) {
		value = value.Mid(5, value.Length() - 6);

		wxString startCell = value.BeforeFirst(':');
		wxString endCell = value.AfterFirst(':');

		char ch1 = startCell[0];
		char ch2 = endCell[0];

		int startCol = ch1 - 65;
		int endCol = ch2 - 65;

		int startRow = wxAtoi(startCell.Mid(1)) - 1;
		int endRow = wxAtoi(endCell.Mid(1)) - 1;

		int sum = 0;
		for (int r = startRow; r <= endRow;r++) {
			for (int c = startCol;c <= endCol;c++) {
				wxString cellValue = grid->GetCellValue(r, c);
				long number;
				if (cellValue.ToLong(&number)) {
					sum += number;
				}
			}
		}

		grid->SetCellValue(row, col, wxString::Format("%d", sum));

	}

	else if (value.StartsWith("=AVERAGE(") && value.EndsWith(")")) {
		value = value.Mid(9, value.Length() - 10);

		wxString startCell = value.BeforeFirst(':');
		wxString endCell = value.AfterFirst(':');

		char ch1 = startCell[0];
		char ch2 = endCell[0];

		int startCol = ch1 - 65;
		int endCol = ch2 - 65;

		int startRow = wxAtoi(startCell.Mid(1)) - 1;
		int endRow = wxAtoi(endCell.Mid(1)) - 1;

		double sum = 0;
		int count = 0;

		for (int r = startRow; r <= endRow; r++) {
			for (int c = startCol; c <= endCol; c++) {
				wxString cellValue = grid->GetCellValue(r, c);
				double number;
				if (cellValue.ToDouble(&number)) {
					sum += number;
					count++;
				}
			}
		}

		if (count > 0) {
			double average = sum / count;
			grid->SetCellValue(row, col, wxString::Format("%.2f", average));
		}
		else {
			grid->SetCellValue(row, col, "ERROR");
		}
	}


}

void MyFrame::OnSave(wxCommandEvent& event) {
	wxFileDialog saveFileDialog(this, "Save CSV file", "", "",
		"CSV files (*.csv)|*.csv", wxFD_SAVE | wxFD_OVERWRITE_PROMPT);

	if (saveFileDialog.ShowModal() == wxID_CANCEL)
		return;

	wxString filePath = saveFileDialog.GetPath();

	wxFile file(filePath, wxFile::write);

	if (!file.IsOpened()) {
		wxMessageBox("Unable to create file", "Error", wxICON_ERROR);
		return;
	}

	int rows = grid->GetNumberRows();
	int cols = grid->GetNumberCols();

	for (int r = 0; r < rows; r++) {
		wxString line;
		for (int c = 0; c < cols; c++) {
			if (c > 0) line += ",";
			line += grid->GetCellValue(r, c);
		}
		line += "\n";
		file.Write(line);
	}

	file.Close();
	wxMessageBox("File saved successfully!", "Success", wxICON_INFORMATION);
}


void MyFrame::OnOpen(wxCommandEvent& event) {
	wxFileDialog openFileDialog(this, "Open CSV file", "", "",
		"CSV files (*.csv)|*.csv", wxFD_OPEN | wxFD_FILE_MUST_EXIST);

	if (openFileDialog.ShowModal() == wxID_CANCEL)
		return;

	wxString filePath = openFileDialog.GetPath();
	wxTextFile file(filePath);

	if (!file.Open()) {
		wxMessageBox("Unable to open file", "Error", wxICON_ERROR);
		return;
	}

	if (grid->GetNumberRows() > 0) {
		grid->ClearGrid();
		grid->DeleteRows(0, grid->GetNumberRows());
	}

	int maxCols = 0;

	while (!file.Eof()) {
		wxString line = file.GetNextLine();
		wxArrayString values = wxStringTokenize(line, ",");
		if (values.GetCount() > maxCols) {
			maxCols = values.GetCount();
		}
	}

	file.Close();


	int currentCols = grid->GetNumberCols();
	if (currentCols < maxCols) {
		grid->AppendCols(maxCols - currentCols);
	}

	wxTextFile file2(filePath);
	if (!file2.Open()) {
		wxMessageBox("Unable to reopen file", "Error", wxICON_ERROR);
		return;
	}

	int r = 0;
	while (!file2.Eof()) {
		wxString line = file2.GetNextLine();
		wxArrayString values = wxStringTokenize(line, ",");

		if (values.GetCount() > 0) {
			if (grid->GetNumberRows() <= r) {
				grid->AppendRows(1);
			}

			for (size_t c = 0; c < values.GetCount(); c++) {
				grid->SetCellValue(r, c, values[c]);
			}
		}
		r++;
	}

	file2.Close();
}

void MyFrame::OnExit(wxCommandEvent& event) {
	Close(true);
}


