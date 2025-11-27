// https://74pbcspn-3001.use2.devtunnels.ms/api/clients/:id'

const API_BASE_URL = 'https://74pbcspn-3002.use2.devtunnels.ms/api/clients';
const API_CLIENTS_URL = `${API_BASE_URL}/`;


export const getClients = async (dataclientes) => {
  try {
    const response = await fetch(API_CLIENTS_URL);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    console.log(data);
    return data;
  } catch (error) {
    console.error('Error fetching clients:', error);
    throw error;
  }
}