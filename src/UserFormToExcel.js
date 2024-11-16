import { saveAs } from "file-saver";
import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import './UserFormToExcel.css'; // Import the external CSS file for styling

const UserFormToExcel = () => {
  const [formData, setFormData] = useState({
    name: "",
    email: "",
    age: "",
    phone: "",
    sapCode: "",
    sapDescription: "",
    hsCode: "",
    remarks: "",
    image: "", // Store image path as a string
    countryCode: "+1", // Default country code (US)
  });

  const [userData, setUserData] = useState([]);
  const [workbook, setWorkbook] = useState(null); // Store the Excel workbook
  const [successMessage, setSuccessMessage] = useState("");
  const [filePath, setFilePath] = useState("");

  useEffect(() => {
    // Initialize an empty workbook on component mount
    const initialWorkbook = XLSX.utils.book_new();
    const initialWorksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(initialWorkbook, initialWorksheet, "User Data");
    setWorkbook(initialWorkbook);
  }, []);

  const handleChange = (event) => {
    const { name, value } = event.target;
    setFormData({ ...formData, [name]: value });
  };

  const handleCountryCodeChange = (event) => {
    setFormData({ ...formData, countryCode: event.target.value });
  };

  const handleSubmit = (event) => {
    event.preventDefault();
    setUserData((prevData) => [...prevData, formData]);
    setFormData({
      name: "",
      email: "",
      age: "",
      phone: "",
      sapCode: "",
      sapDescription: "",
      hsCode: "",
      remarks: "",
      image: "", // Clear image path
      countryCode: "+1",
    }); // Clear form
  };

  const handleImageUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      // Assuming the image is uploaded to the server and we receive the file path
      const imagePath = `/uploads/images/${file.name}`; // Example path
      setFormData({ ...formData, image: imagePath }); // Store image path in state
    }
  };

  const exportToExcel = () => {
    if (userData.length === 0) {
      alert("No data to export.");
      return;
    }

    let sheetName = "User Data";
    if (workbook.SheetNames.includes(sheetName)) {
      workbook.SheetNames = workbook.SheetNames.filter((name) => name !== sheetName);
    }

    const worksheet = XLSX.utils.json_to_sheet(userData);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "binary" });
    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    const fileName = "User_Data.xlsx";
    saveAs(blob, fileName);

    setFilePath(`E:/Abir File/${fileName}`);
    setSuccessMessage("Data saved to Excel successfully!");
  };

  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
      view[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
  };

  return (
    <div className="form-container">
      <h1>User Form</h1>
      <form onSubmit={handleSubmit} className="form">
        <div className="input-group">
          <label>Name:</label>
          <input
            type="text"
            name="name"
            value={formData.name}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>Email:</label>
          <input
            type="email"
            name="email"
            value={formData.email}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>Age:</label>
          <input
            type="number"
            name="age"
            value={formData.age}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>Phone:</label>
          <select
            name="countryCode"
            value={formData.countryCode}
            onChange={handleCountryCodeChange}
            className="input-field country-code-select"
          >
            <option value="+1">+1 (USA)</option>
            <option value="+44">+44 (UK)</option>
            <option value="+91">+91 (India)</option>
          </select>
          <input
            type="text"
            name="phone"
            value={formData.phone}
            onChange={handleChange}
            required
            className="input-field phone-input"
            placeholder="Phone Number"
          />
        </div>

        <div className="input-group">
          <label>SAP Code:</label>
          <input
            type="number"
            name="sapCode"
            value={formData.sapCode}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>SAP Description:</label>
          <input
            type="text"
            name="sapDescription"
            value={formData.sapDescription}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>HS Code:</label>
          <input
            type="number"
            name="hsCode"
            value={formData.hsCode}
            onChange={handleChange}
            required
            className="input-field"
          />
        </div>

        <div className="input-group">
          <label>Remarks:</label>
          <textarea
            name="remarks"
            value={formData.remarks}
            onChange={handleChange}
            required
            className="input-field remarks-textarea"
          />
        </div>

        <div className="input-group">
          <label>Image Upload:</label>
          <input
            type="file"
            name="image"
            onChange={handleImageUpload}
            accept="image/*"
            className="input-field"
          />
        </div>

        <div className="button-group">
          <button type="submit" className="submit-button">Add User</button>
          <button type="button" onClick={exportToExcel} className="export-button">Export to Excel</button>
        </div>
      </form>

      {successMessage && <p className="success-message">{successMessage}</p>}
      {filePath && (
        <p className="file-path">
          File saved at: <strong>{filePath}</strong>
        </p>
      )}

      {userData.length > 0 && (
        <div className="user-data-table">
          <h3>Collected User Data:</h3>
          <table className="table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Email</th>
                <th>Age</th>
                <th>Phone</th>
                <th>SAP Code</th>
                <th>SAP Description</th>
                <th>HS Code</th>
                <th>Remarks</th>
                <th>Image Path</th>
              </tr>
            </thead>
            <tbody>
              {userData.map((user, index) => (
                <tr key={index}>
                  <td>{user.name}</td>
                  <td>{user.email}</td>
                  <td>{user.age}</td>
                  <td>{user.countryCode + " " + user.phone}</td>
                  <td>{user.sapCode}</td>
                  <td>{user.sapDescription}</td>
                  <td>{user.hsCode}</td>
                  <td>{user.remarks}</td>
                  <td>{user.image ? user.image : "No image"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default UserFormToExcel;
