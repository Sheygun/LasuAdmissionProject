#include <iostream>
#include "libxl.h"
#include <iomanip>
using namespace std;
using namespace libxl;
struct Program_data
{
    int no_of_admitted_students = 0;
    int no_of_notadmitted_students = 0;
};

class StudentAdmissionCalculator
{
    public:
        void initialize_program(StudentAdmissionCalculator instant);
        void open_excel_sheet(StudentAdmissionCalculator instant);
        void close_excel_sheet(StudentAdmissionCalculator class_instant,Book* instant,string error);
        void process_admission_details(StudentAdmissionCalculator instant,Sheet* excel_sheet,Book* excel_book);
        void process_admission_status(StudentAdmissionCalculator instant,Book* excel_book,Sheet* excel_sheet,int student_row,double aggregate);
        int process_sitting(int no_of_sitting);
        int process_grades(string grade);
        int student_row_start = 4;
        int column_start_count = 3;
        int no_of_students = 0;
        int admission_cut_off_mark = 60; //admission cut off for law students
};

void StudentAdmissionCalculator::process_admission_status(StudentAdmissionCalculator instant,Book* excel_book,Sheet* excel_sheet,int student_row,double student_aggregate)
{
    int student_admitted_count = (student_row - instant.column_start_count);
    if(student_admitted_count <= 50){
        if(student_aggregate >= instant.admission_cut_off_mark){
            Font* success_font = excel_book->addFont();
            success_font->setColor(COLOR_BLUE);
            success_font->setBold(true);
            Format* bold_success_Format = excel_book->addFormat();
            bold_success_Format->setFont(success_font);
            excel_sheet->writeStr(student_row,11, "ADMITTED",bold_success_Format);
        }else{
            Font* failed_font = excel_book->addFont();
            failed_font->setColor(COLOR_RED);
            failed_font->setBold(true);
            Format* bold_failed_Format = excel_book->addFormat();
            bold_failed_Format->setFont(failed_font);
            excel_sheet->writeStr(student_row,11, "NOT ADMITTED",bold_failed_Format);
        }
    }else{
        Font* failed_font = excel_book->addFont();
        failed_font->setColor(COLOR_RED);
        failed_font->setBold(true);
        Format* bold_failed_Format = excel_book->addFormat();
        bold_failed_Format->setFont(failed_font);
        excel_sheet->writeStr(student_row,11, "NOT ADMITTED.",bold_failed_Format);
    }
}

void StudentAdmissionCalculator::process_admission_details(StudentAdmissionCalculator instant,Sheet* excel_sheet,Book* excel_book)
{
    for(int row_count = 0; row_count < instant.no_of_students; row_count++){
        int student_row = student_row_start + row_count;
        int sitting_mark = 0, total_grade_score = 0;
        int utme_point = 0,putme_point = 0,olevel_point = 0;
        double aggregate = 0;
        if(excel_sheet)
        {
            sitting_mark = instant.process_sitting(excel_sheet->readNum(student_row,2));
            utme_point = (excel_sheet->readNum(student_row,8) * 0.0625);
            putme_point = (excel_sheet->readNum(student_row,9) * 0.5);
            //loop for waec grades
            for(int column_count = 0; column_count < 5; column_count++)
            {
                int grade_col = column_start_count + column_count;
                total_grade_score += instant.process_grades(excel_sheet->readStr(student_row,grade_col));
            }
            olevel_point = 0.5 * (total_grade_score + sitting_mark);
            aggregate = olevel_point + putme_point + utme_point;
            excel_sheet->writeNum(student_row,10,aggregate);
            instant.process_admission_status(instant,excel_book,excel_sheet,student_row,aggregate);
        }
    }
    instant.close_excel_sheet(instant,excel_book,"****\nSuccessful");
}

int StudentAdmissionCalculator::process_grades(string grade)
{
    if (grade=="A1"){
        return 8;
    }else if (grade=="B2"){
        return 7;
    }else if (grade=="B3"){
        return 6;
    }else if (grade=="C4"){
        return 5;
    }else if (grade=="C5"){
        return 4;
    }else if (grade=="C6"){
        return 3;
    }else{
        return 0;
    }
}

int StudentAdmissionCalculator::process_sitting(int sitting_count)
{
    switch(sitting_count)
    {
        case 1:
            return 10;
            break;
        case 2:
            return 5;
            break;
    }
}

void StudentAdmissionCalculator::open_excel_sheet(StudentAdmissionCalculator instant)
{
    Book* excel_book = xlCreateBook();
    excel_book->load("lawschool_admission_record.xls");
    if(excel_book)
    {
        Sheet* excel_sheet = excel_book->getSheet(0);
        if(excel_sheet)
        {
            instant.process_admission_details(instant,excel_sheet,excel_book);
        }else{
            cout << "Unable to open excel sheet" << endl;
        }
    }else{
        cout << "Unable to open excel document" << endl;
    }
}

void StudentAdmissionCalculator::close_excel_sheet(StudentAdmissionCalculator instant,Book* excel_book,string msg)
{
    excel_book->save("lawschool_admission_record.xls");
    excel_book->release();
    cout << "\n" << msg << endl;
}

void StudentAdmissionCalculator::initialize_program(StudentAdmissionCalculator class_instant)
{
    cout << "\n\n";
    cout << "Welcome to LASU Student Admission(Law Students)" << endl;
    cout << "\n";
    cout << ".... Loading Excel sheet" << endl;
    cout << "Enter no of Students : ";
    cin >> class_instant.no_of_students;
    if(class_instant.no_of_students > 0)
    {
        cout << "Using 60% as cut off mark for Law Students....." << endl;
        class_instant.open_excel_sheet(class_instant);
    }else
    {
        cout << "You entered zero number of students.." << endl;
        cout << "Exiting Application ........" << endl;
    }
}

int main()
{
    StudentAdmissionCalculator class_instant;
    class_instant.initialize_program(class_instant);
    return 0;
}
