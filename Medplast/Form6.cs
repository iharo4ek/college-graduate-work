﻿using System;
                default: {
                        this.mode = Modes.READWRITE;
                        break;
            string query = "select id_employee,(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [driver] from employees;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, dataBase.getSrc());
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "id_employee";
            comboBox1.DisplayMember = "driver";
        }
            textBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                if (textBox1.Text.Length < 3) {
                    MessageBox.Show("Слишком короткое название марки", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string query = $"Insert into cars values (N'{maskedTextBox1.Text}', {comboBox1.SelectedValue}, N'{textBox1.Text}', '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}');";
                SqlCommand sqlCommand = new SqlCommand(query, sql);
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Данные успешно добавлены", "SUCCESS", MessageBoxButtons.OK);
                cars();
            } catch (Exception ex) {
                if (textBox1.Text.Length < 3) {
                    MessageBox.Show("Слишком короткое название марки", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                string query = $"Update cars Set numberOfTheCar = N'{maskedTextBox1.Text}', id_driver = {comboBox1.SelectedValue},carBrand = N'{textBox1.Text}', dateOfpurchase = '{dateTimePicker1.Value.ToString("yyyy-MM-dd")}'  where id_car = {dataGridView1.CurrentRow.Cells[0].Value}";
                SqlCommand com = new SqlCommand(query, sql);
                com.ExecuteNonQuery();
                cars();
                MessageBox.Show("Данные успешно изменены", "SUCCESS", MessageBoxButtons.OK);
        }
                exelApp.Cells[1, 1] = "машины ОАО Медпласт";
                for (int i = 0; i < dataGridView1.RowCount; i++) {
        private void radioButton2_CheckedChanged(object sender, EventArgs e) {
            if (radioButton2.Checked == true) {
                dateTimePicker2.Visible = true;
                dateTimePicker3.Visible = true;
                button4.Enabled = true;
            } else {
                dateTimePicker2.Visible = false;
                dateTimePicker3.Visible = false;
            }
        }
        private void button4_Click(object sender, EventArgs e) {
            if (radioButton1.Checked) {
                if (textBox3.Text.Length == 0) {
                    MessageBox.Show($"Для фильтрации необходимо заполнить данные", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } else {
                    filter = $"select id_car, numberOfTheCar as  [гос. номер],(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [водитель],carBrand as [марка], dateOfpurchase as [дата покупки] from cars inner join employees on cars.id_driver = employees.id_employee where carBrand like N'{textBox3.Text}';";
                }
            }
            if (radioButton2.Checked) {
                filter = $"select id_car, numberOfTheCar as  [гос. номер],(employeeSurname + ' ' + employeeName + ' ' + employeePatronymic) as [водитель],carBrand as [марка], dateOfpurchase as [дата покупки] from cars inner join employees on cars.id_driver = employees.id_employee where dateOfpurchase >= '{dateTimePicker2.Value.ToString("yyyy-MM-dd")}' and  dateOfpurchase <= '{dateTimePicker3.Value.ToString("yyyy-MM-dd")}';";
            }
            DataTable dt = GetData(filter);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AutoResizeColumns();
            button6.Enabled = true;
        }
        private void textBox3_TextChanged(object sender, EventArgs e) {
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e) {
            if (radioButton1.Checked == true) {
                textBox3.Visible = true;
                button4.Enabled = true;
            } else {
                textBox3.Visible = false;
            }
        }
        private void button6_Click(object sender, EventArgs e) {
            cars();
            button6.Enabled = false;
        }
        private void Form6_Shown(object sender, EventArgs e) {
            sql = dataBase.getConnection();
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {

        }
    }