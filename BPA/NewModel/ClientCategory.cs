using BPA.NewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NM = BPA.NewModel;

namespace BPA.Modules
{
    class ClientCategory
    {
        public string CustomerStatus;
        public string ChannelType;
        public string Mag;

        public ClientCategory GetCategoryFromClient(NM.ClientItem client)
        {
            if (client ==null)
                return null;

            ClientCategory clientCategory = new ClientCategory
            { 
                CustomerStatus = client.CustomerStatus,
                ChannelType = client.ChannelType,
                Mag = client.Mag
            };

            return clientCategory;
        }

        public List<ClientCategory> GetCategoryListFromClients(NM.ClientTable clients)
        {
            if (clients == null)
                return null;

            List<ClientCategory> clientCategories = new List<ClientCategory>();

            foreach(ClientItem client in clients)
            {
                ClientCategory tmp = clientCategories.Find(x => x.ChannelType == client.ChannelType && x.CustomerStatus == client.CustomerStatus);
                if (tmp != null) continue;

                clientCategories.Add(GetCategoryFromClient(client));
            }

            return clientCategories;
        }

        public List<ClientCategory> GetCategoryListFromClients(List<NM.ClientItem> clients)
        {
            if (clients == null)
                return null;

            List<ClientCategory> clientCategories = new List<ClientCategory>();

            foreach (ClientItem client in clients)
                clientCategories.Add(GetCategoryFromClient(client));

            return clientCategories;
        }
    }
}
