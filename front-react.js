import React, { useState } from 'react';
import axios from 'axios';

function App() {
  const [txtFile, setTxtFile] = useState(null);
  const [imageFile, setImageFile] = useState(null);

  const handleTxtFileChange = (e) => {
    setTxtFile(e.target.files[0]);
  };

  const handleImageFileChange = (e) => {
    setImageFile(e.target.files[0]);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    const formData = new FormData();
    formData.append('txt_file', txtFile);
    formData.append('image_file', imageFile);

    try {
      const response = await axios.post('/upload', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'presentation.pptx');
      document.body.appendChild(link);
      link.click();
    } catch (error) {
      console.error('Error uploading files', error);
    }
  };

  return (
    <div>
      <form onSubmit={handleSubmit}>
        <input type="file" onChange={handleTxtFileChange} />
        <input type="file" onChange={handleImageFileChange} />
        <button type="submit">Generate Presentation</button>
      </form>
    </div>
  );
}

export default App;