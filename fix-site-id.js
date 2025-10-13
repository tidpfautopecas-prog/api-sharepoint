
import fetch from 'node-fetch';
import dotenv from 'dotenv';

dotenv.config();

console.log('🔍 Descobrindo o SITE_ID e bibliotecas corretas do SharePoint...');

async function getAccessToken() {
  try {
    const params = new URLSearchParams();
    params.append('client_id', process.env.CLIENT_ID);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('client_secret', process.env.CLIENT_SECRET);
    params.append('grant_type', 'client_credentials');

    const res = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      body: params,
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const data = await res.json();
    
    if (!data.access_token) {
      throw new Error(`Erro na autenticação: ${data.error_description || data.error}`);
    }

    return data.access_token;
  } catch (error) {
    console.error('❌ Erro ao obter token:', error.message);
    throw error;
  }
}

async function findCorrectSiteAndLibraries() {
  try {
    const accessToken = await getAccessToken();
    
    console.log('✅ Token obtido com sucesso');
    console.log('🔍 Testando site GLB-FS...');

    // Testar o site atual
    const currentSiteId = process.env.SITE_ID;
    console.log(`📍 Site ID atual: ${currentSiteId}`);
    
    const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}`, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!siteResponse.ok) {
      console.log('❌ Site atual não funciona, buscando alternativas...');
      
      // Buscar por sites com GLB
      const searchResponse = await fetch('https://graph.microsoft.com/v1.0/sites?search=GLB', {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      
      if (searchResponse.ok) {
        const searchData = await searchResponse.json();
        console.log('\n🔍 SITES ENCONTRADOS COM "GLB":');
        console.log('='.repeat(80));
        
        searchData.value.forEach((site, index) => {
          console.log(`${index + 1}. Nome: ${site.displayName}`);
          console.log(`   URL: ${site.webUrl}`);
          console.log(`   ID: ${site.id}`);
          console.log('-'.repeat(40));
        });
        
        // Tentar o primeiro site encontrado
        if (searchData.value.length > 0) {
          const firstSite = searchData.value[0];
          console.log(`\n🎯 TESTANDO PRIMEIRO SITE: ${firstSite.displayName}`);
          console.log(`📝 NOVO SITE_ID: ${firstSite.id}`);
          
          // Testar bibliotecas deste site
          await testSiteLibraries(accessToken, firstSite.id, firstSite.displayName);
        }
      }
      return;
    }

    const siteData = await siteResponse.json();
    console.log(`✅ Site acessível: ${siteData.displayName}`);
    console.log(`📍 URL: ${siteData.webUrl}`);
    
    // Listar todas as bibliotecas/drives disponíveis
    await testSiteLibraries(accessToken, currentSiteId, siteData.displayName);

  } catch (error) {
    console.error('❌ Erro ao buscar sites:', error.message);
  }
}

async function testSiteLibraries(accessToken, siteId, siteName) {
  try {
    console.log(`\n📚 BIBLIOTECAS DISPONÍVEIS NO SITE: ${siteName}`);
    console.log('='.repeat(80));
    
    const librariesResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    
    if (!librariesResponse.ok) {
      throw new Error(`Erro ao listar bibliotecas: ${librariesResponse.status}`);
    }
    
    const librariesData = await librariesResponse.json();
    
    if (librariesData.value.length === 0) {
      console.log('❌ Nenhuma biblioteca encontrada');
      return;
    }
    
    librariesData.value.forEach((library, index) => {
      console.log(`${index + 1}. Nome: ${library.name}`);
      console.log(`   Descrição: ${library.description || 'N/A'}`);
      console.log(`   ID: ${library.id}`);
      console.log(`   Tipo: ${library.driveType}`);
      console.log('-'.repeat(40));
    });
    
    // Tentar encontrar a biblioteca correta
    const possibleLibraries = librariesData.value.filter(lib => 
      lib.name.toLowerCase().includes('document') || 
      lib.name.toLowerCase().includes('shared') ||
      lib.name.toLowerCase().includes('compartilhad') ||
      lib.driveType === 'documentLibrary'
    );
    
    if (possibleLibraries.length > 0) {
      console.log('\n🎯 BIBLIOTECAS CANDIDATAS PARA DOCUMENTOS:');
      console.log('='.repeat(80));
      
      for (const library of possibleLibraries) {
        console.log(`📁 Testando biblioteca: ${library.name}`);
        
        // Testar acesso à biblioteca
        try {
          const testUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${library.id}/root/children`;
          const testResponse = await fetch(testUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          
          if (testResponse.ok) {
            const children = await testResponse.json();
            console.log(`   ✅ Acessível - ${children.value.length} itens encontrados`);
            
            // Listar pastas principais
            const folders = children.value.filter(item => item.folder);
            if (folders.length > 0) {
              console.log(`   📂 Pastas principais:`);
              folders.forEach(folder => {
                console.log(`      - ${folder.name}`);
              });
            }
            
            // Verificar se já existe pasta Laudos
            const laudosFolder = children.value.find(item => 
              item.folder && item.name.toLowerCase().includes('laudo')
            );
            
            if (laudosFolder) {
              console.log(`   🎯 PASTA LAUDOS ENCONTRADA: ${laudosFolder.name}`);
            }
            
            console.log(`\n📝 CONFIGURAÇÃO RECOMENDADA PARA .env:`);
            console.log(`SITE_ID=${siteId}`);
            console.log(`LIBRARY_NAME=${library.name}`);
            console.log(`FOLDER_PATH=Laudos`);
            console.log('='.repeat(80));
            
          } else {
            console.log(`   ❌ Não acessível - Status: ${testResponse.status}`);
          }
        } catch (error) {
          console.log(`   ❌ Erro ao testar: ${error.message}`);
        }
        
        console.log('-'.repeat(40));
      }
    }
    
    // Testar criação de pasta Laudos na primeira biblioteca válida
    if (possibleLibraries.length > 0) {
      const mainLibrary = possibleLibraries[0];
      console.log(`\n🧪 TESTANDO CRIAÇÃO DE PASTA LAUDOS EM: ${mainLibrary.name}`);
      
      try {
        const createFolderUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${mainLibrary.id}/root/children`;
        
        const createResponse = await fetch(createFolderUrl, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            name: 'Laudos',
            folder: {},
            '@microsoft.graph.conflictBehavior': 'rename'
          })
        });
        
        if (createResponse.ok) {
          const result = await createResponse.json();
          console.log(`✅ Pasta Laudos criada com sucesso!`);
          console.log(`📍 URL: ${result.webUrl}`);
        } else if (createResponse.status === 409) {
          console.log(`✅ Pasta Laudos já existe!`);
        } else {
          const errorText = await createResponse.text();
          console.log(`❌ Erro ao criar pasta: ${createResponse.status} - ${errorText}`);
        }
      } catch (error) {
        console.log(`❌ Erro ao testar criação de pasta: ${error.message}`);
      }
    }
    
  } catch (error) {
    console.error('❌ Erro ao listar bibliotecas:', error.message);
  }
}

// Executar
findCorrectSiteAndLibraries();
